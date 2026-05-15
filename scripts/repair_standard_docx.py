#!/usr/bin/env python3
"""Repair a DOCX according to the standard Word formatting checklist."""

from __future__ import annotations

import argparse
import re
from collections import Counter
from pathlib import Path

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from docx.text.paragraph import Paragraph
except ImportError as exc:  # pragma: no cover
    raise SystemExit("missing dependency: python-docx. Install with: pip install python-docx") from exc

from audit_standard_docx import (
    DEFAULT_OUTPUT_DIR,
    UNICODE_BULLETS,
    UNSAFE_FONTS,
    collect_issues,
    heading_level,
    is_body,
    is_heading,
    looks_fake_heading,
    paragraph_text,
    run_font_names,
    run_size_pt,
)


WESTERN_FONT = "Arial"
EAST_ASIA_FONT = "微软雅黑"


def set_font_east_asia(run, western: str = WESTERN_FONT, east_asia: str = EAST_ASIA_FONT) -> None:
    run.font.name = western
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr, value in (
        ("w:ascii", western),
        ("w:hAnsi", western),
        ("w:cs", western),
        ("w:eastAsia", east_asia),
    ):
        rfonts.set(qn(attr), value)


def safe_set_style(paragraph, style_name: str) -> None:
    try:
        paragraph.style = style_name
    except KeyError:
        paragraph.style = "Normal"


def rewrite_paragraph(paragraph, text: str, style_name: str | None = None) -> None:
    paragraph.clear()
    if style_name:
        safe_set_style(paragraph, style_name)
    run = paragraph.add_run(text)
    set_font_east_asia(run)
    run.font.size = Pt(11)


def remove_paragraph(paragraph) -> None:
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def insert_paragraph_after(paragraph, text: str) -> Paragraph:
    new_element = OxmlElement("w:p")
    paragraph._p.addnext(new_element)
    new_paragraph = Paragraph(new_element, paragraph._parent)
    try:
        new_paragraph.style = paragraph.style
    except Exception:
        safe_set_style(new_paragraph, "Normal")
    run = new_paragraph.add_run(text)
    set_font_east_asia(run)
    run.font.size = Pt(11)
    return new_paragraph


def paragraph_has_hard_break(paragraph) -> bool:
    for br in paragraph._p.xpath(".//w:br"):
        if br.get(qn("w:type")) not in {"page", "column"}:
            return True
    return "\v" in paragraph.text or "\n" in paragraph.text


def repair_unicode_bullet(paragraph) -> bool:
    text = paragraph_text(paragraph)
    stripped = text.lstrip()
    if not stripped.startswith(UNICODE_BULLETS):
        return False
    leading_spaces = len(text) - len(stripped)
    cleaned = stripped.lstrip("".join(UNICODE_BULLETS)).strip()
    rewrite_paragraph(paragraph, " " * leading_spaces + cleaned, "List Bullet")
    return True


def repair_fake_heading(paragraph) -> bool:
    if not looks_fake_heading(paragraph):
        return False
    max_size = max((run_size_pt(run) or 0 for run in paragraph.runs), default=0)
    if max_size >= 18:
        style = "Heading 1"
    elif max_size >= 14:
        style = "Heading 2"
    elif max_size >= 12:
        style = "Heading 3"
    else:
        style = "Heading 4"
    safe_set_style(paragraph, style)
    for run in paragraph.runs:
        set_font_east_asia(run)
    return True


def repair_unsafe_fonts(paragraph) -> int:
    count = 0
    for run in paragraph.runs:
        if run_font_names(run) & UNSAFE_FONTS:
            set_font_east_asia(run)
            count += 1
    return count


def repair_hard_break(paragraph) -> bool:
    if not paragraph_has_hard_break(paragraph):
        return False
    parts = [part.strip() for part in re.split(r"[\n\v]+", paragraph.text) if part.strip()]
    if len(parts) <= 1:
        return False
    rewrite_paragraph(paragraph, parts[0])
    cursor = paragraph
    for part in parts[1:]:
        cursor = insert_paragraph_after(cursor, part)
    return True


def repair_body_font_sizes(doc) -> int:
    body_sizes: set[float] = set()
    for paragraph in doc.paragraphs:
        if not is_body(paragraph):
            continue
        for run in paragraph.runs:
            size = run_size_pt(run)
            if size is not None:
                body_sizes.add(size)
    if len(body_sizes) <= 2:
        return 0
    changed = 0
    for paragraph in doc.paragraphs:
        if not is_body(paragraph):
            continue
        for run in paragraph.runs:
            run.font.size = Pt(11)
            set_font_east_asia(run)
            changed += 1
    return changed


def set_table_width(table, width_emu: int) -> None:
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(int(width_emu / 635)))


def repair_tables(doc) -> int:
    if not doc.tables:
        return 0
    section = doc.sections[0]
    content_width = int(section.page_width - section.left_margin - section.right_margin)
    changed = 0
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_width(table, content_width)
        cols = len(table.columns) or 1
        col_width = int(content_width / cols)
        for row in table.rows:
            for cell in row.cells:
                cell.width = col_width
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        set_font_east_asia(run)
        changed += 1
    return changed


def repair_empty_paragraphs(doc) -> int:
    removed = 0
    empty_run = 0
    for paragraph in list(doc.paragraphs):
        if paragraph_text(paragraph):
            empty_run = 0
            continue
        empty_run += 1
        if empty_run >= 2:
            remove_paragraph(paragraph)
            removed += 1
    return removed


def output_paths(input_path: Path, output: Path | None) -> tuple[Path, Path]:
    if output:
        repaired = output
    else:
        repaired = input_path.with_name(f"{input_path.stem}_repaired.docx")
    if repaired.resolve() == input_path.resolve():
        raise SystemExit("repair output must not overwrite the original file")
    summary = repaired.with_name(f"{repaired.stem.replace('_repaired', '')}_repair_summary.md")
    return repaired, summary


def write_summary(input_path: Path, repaired_path: Path, summary_path: Path, before, after, fixes: Counter) -> None:
    unresolved = [issue for issue in after if not issue.auto_fixable]
    with summary_path.open("w", encoding="utf-8") as fh:
        fh.write("# Word 格式修复摘要\n\n")
        fh.write(f"- 原文件：`{input_path}`\n")
        fh.write(f"- 修复后文件：`{repaired_path}`\n")
        fh.write(f"- 修复前问题数：{len(before)}\n")
        fh.write(f"- 修复后剩余问题数：{len(after)}\n\n")
        fh.write("## 自动修复\n\n")
        if fixes:
            for rule_id, count in sorted(fixes.items()):
                fh.write(f"- `{rule_id}`：{count} 处\n")
        else:
            fh.write("- 未执行自动修复。\n")
        fh.write("\n## 仍需人工确认\n\n")
        if unresolved:
            for issue in unresolved:
                fh.write(f"- `{issue.rule_id}` {issue.location}：{issue.message} {issue.action}\n")
        else:
            fh.write("- 无。\n")


def repair(input_path: Path, output: Path | None = None) -> tuple[Path, Path]:
    if not input_path.exists():
        raise SystemExit(f"input not found: {input_path}")
    if input_path.suffix.lower() != ".docx":
        raise SystemExit("only .docx is supported")

    before = collect_issues(input_path)
    before_rule_ids = {issue.rule_id for issue in before}
    doc = Document(str(input_path))
    fixes: Counter = Counter()

    for paragraph in list(doc.paragraphs):
        if repair_unicode_bullet(paragraph):
            fixes["E001"] += 1
        if repair_fake_heading(paragraph):
            fixes["E002"] += 1
        unsafe_count = repair_unsafe_fonts(paragraph)
        if unsafe_count:
            fixes["E003"] += unsafe_count
        if repair_hard_break(paragraph):
            fixes["E004"] += 1

    size_changes = repair_body_font_sizes(doc)
    if size_changes:
        fixes["W001"] += size_changes

    table_changes = repair_tables(doc)
    if table_changes and "W002" in before_rule_ids:
        fixes["W002"] += table_changes

    empty_changes = repair_empty_paragraphs(doc)
    if empty_changes:
        fixes["W003"] += empty_changes

    repaired_path, summary_path = output_paths(input_path, output)
    repaired_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(repaired_path)
    after = collect_issues(repaired_path)
    write_summary(input_path, repaired_path, summary_path, before, after, fixes)
    return repaired_path, summary_path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", type=Path, help="Optional repaired .docx output path")
    args = parser.parse_args()
    repaired, summary = repair(args.input.expanduser(), args.output.expanduser() if args.output else None)
    print(repaired)
    print(summary)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
