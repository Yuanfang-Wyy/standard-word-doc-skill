#!/usr/bin/env python3
"""Audit a DOCX against the standard Word material rules."""

from __future__ import annotations

import argparse
import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
DEFAULT_OUTPUT_DIR = Path.home() / "Documents" / "AI-Stack-Outputs" / "word-docs"


def qn(name: str) -> str:
    prefix, tag = name.split(":", 1)
    if prefix != "w":
        raise ValueError(name)
    return f"{{{W_NS}}}{tag}"


def para_text(p: ET.Element) -> str:
    return "".join(t.text or "" for t in p.findall(".//w:t", NS)).strip()


def style_id(p: ET.Element) -> str:
    el = p.find("./w:pPr/w:pStyle", NS)
    return el.attrib.get(qn("w:val"), "") if el is not None else ""


def has_field(root: ET.Element, keyword: str) -> bool:
    for instr in root.findall(".//w:instrText", NS):
        if instr.text and keyword.upper() in instr.text.upper():
            return True
    return False


def compact_text(text: str) -> str:
    return re.sub(r"\s+", "", text)


def style_name_map(styles_xml: str) -> dict[str, str]:
    if not styles_xml:
        return {}
    try:
        root = ET.fromstring(styles_xml.encode("utf-8"))
    except ET.ParseError:
        return {}
    result: dict[str, str] = {}
    for style in root.findall(".//w:style", NS):
        sid = style.attrib.get(qn("w:styleId"), "")
        name_el = style.find("w:name", NS)
        if sid and name_el is not None:
            result[sid] = name_el.attrib.get(qn("w:val"), "")
    return result


def heading_level(style: str, names: dict[str, str]) -> int | None:
    name = names.get(style, "").lower()
    candidates = {
        "2": 1,
        "3": 2,
        "4": 3,
        "5": 4,
    }
    if style in candidates:
        return candidates[style]
    match = re.search(r"heading\s*([1-4])", name)
    if match:
        return int(match.group(1))
    match = re.search(r"标题\s*([1-4一二三四])", name)
    if match:
        raw = match.group(1)
        return {"一": 1, "二": 2, "三": 3, "四": 4}.get(raw, int(raw) if raw.isdigit() else 0) or None
    return None


def style_exists(styles_xml: str, names: dict[str, str], style_id: str, expected_name: str) -> bool:
    if f'w:styleId="{style_id}"' in styles_xml:
        return True
    expected = expected_name.lower()
    return any(expected in name.lower() for name in names.values())


def bullet_numbering_issues(numbering_xml: str) -> list[str]:
    if not numbering_xml:
        return []
    try:
        root = ET.fromstring(numbering_xml.encode("utf-8"))
    except ET.ParseError:
        return []

    issues: list[str] = []
    unsafe_fonts = {"symbol", "wingdings", "wingdings 2", "wingdings 3"}
    for level in root.findall(".//w:lvl", NS):
        num_fmt = level.find("w:numFmt", NS)
        level_text = level.find("w:lvlText", NS)
        is_bullet = (
            num_fmt is not None
            and num_fmt.attrib.get(qn("w:val")) == "bullet"
            or level_text is not None
            and level_text.attrib.get(qn("w:val")) in {"●", "○", "•", "·"}
        )
        if not is_bullet:
            continue
        fonts = level.find("w:rPr/w:rFonts", NS)
        font_values = [value.lower() for value in (fonts.attrib.values() if fonts is not None else [])]
        if any(value in unsafe_fonts for value in font_values):
            issues.append("项目符号编号使用 Symbol/Wingdings 字体，WPS 中可能显示为乱码")
    return issues


def audit(path: Path) -> dict[str, object]:
    if not path.exists():
        raise SystemExit(f"input not found: {path}")
    if path.suffix.lower() != ".docx":
        raise SystemExit("only .docx is supported in v1")
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        document = ET.fromstring(z.read("word/document.xml"))
        styles = z.read("word/styles.xml").decode("utf-8", errors="ignore") if "word/styles.xml" in names else ""
        numbering_xml = z.read("word/numbering.xml").decode("utf-8", errors="ignore") if "word/numbering.xml" in names else ""
        numbering = bool(numbering_xml)
        footers = [n for n in names if n.startswith("word/footer") and n.endswith(".xml")]

    names_by_id = style_name_map(styles)
    paragraphs = document.findall(".//w:p", NS)
    texts = [para_text(p) for p in paragraphs]
    styled = [(style_id(p), para_text(p)) for p in paragraphs if para_text(p)]
    style_counts = Counter(s for s, _ in styled)
    headings = [(heading_level(s, names_by_id), s, t) for s, t in styled if heading_level(s, names_by_id)]
    findings: list[tuple[str, str]] = []

    joined = "\n".join(texts[:60])
    compact_joined = compact_text(joined)
    required_cover_fields = {
        "编制单位": ["编制单位"],
        "文件密级": ["文件密级"],
        "版本": ["版本", "版本号"],
        "日期": ["日期", "编制日期"],
    }
    for label, aliases in required_cover_fields.items():
        if not any(compact_text(alias) in compact_joined for alias in aliases):
            findings.append(("high", f"封面缺少字段：{label}"))

    if not headings:
        findings.append(("high", "未识别到标准标题样式 heading 1-4"))

    prev_level = 0
    for level, _style, text in headings:
        assert level is not None
        if prev_level and level > prev_level + 1:
            findings.append(("medium", f"标题层级跳级：{text}"))
        prev_level = level

    if not has_field(document, "TOC") and "目  录" not in "\n".join(texts[:80]):
        findings.append(("medium", "未识别到目录或 TOC 字段"))

    if not numbering:
        findings.append(("high", "缺少 numbering.xml，自动编号可能不可用"))
    for issue in bullet_numbering_issues(numbering_xml):
        findings.append(("medium", issue))

    if not footers:
        findings.append(("medium", "未发现页脚文件，页码可能缺失"))

    required_styles = {
        "2": "一级标题",
        "3": "二级标题",
        "4": "三级标题",
        "19": "Body Ref",
        "20": "Cover Meta Ref",
        "21": "Table Text Ref",
    }
    for sid, label in required_styles.items():
        if not style_exists(styles, names_by_id, sid, label):
            findings.append(("low", f"与当前标准模板样式不完全一致：{label}"))

    has_body_style = style_counts.get("19", 0) > 0 or any("body ref" in names_by_id.get(s, "").lower() for s in style_counts)
    if not has_body_style:
        findings.append(("low", "未识别到 Body Ref 正文样式，可能使用了其他正文样式"))

    return {
        "path": str(path),
        "paragraph_count": len([t for t in texts if t]),
        "heading_count": len(headings),
        "style_counts": dict(style_counts),
        "has_toc_field": has_field(document, "TOC"),
        "has_numbering": numbering,
        "footer_count": len(footers),
        "headings": headings[:80],
        "findings": findings,
    }


def write_report(result: dict[str, object], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    findings = result["findings"]
    headings = result["headings"]
    with output.open("w", encoding="utf-8") as f:
        f.write(f"# Word 标准规范审计报告\n\n")
        f.write(f"- 文件：`{result['path']}`\n")
        f.write(f"- 段落数：{result['paragraph_count']}\n")
        f.write(f"- 标题数：{result['heading_count']}\n")
        f.write(f"- TOC 字段：{'yes' if result['has_toc_field'] else 'no'}\n")
        f.write(f"- 自动编号文件：{'yes' if result['has_numbering'] else 'no'}\n")
        f.write(f"- 页脚文件数：{result['footer_count']}\n\n")
        f.write("## 发现问题\n\n")
        if findings:
            for severity, message in findings:
                f.write(f"- [{severity}] {message}\n")
        else:
            f.write("- 未发现高优先级结构问题。\n")
        f.write("\n## 标题预览\n\n")
        for level, _style, text in headings:
            f.write(f"- H{level} {text}\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    result = audit(args.input)
    output = args.output or DEFAULT_OUTPUT_DIR / f"{args.input.stem}_audit.md"
    write_report(result, output)
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
