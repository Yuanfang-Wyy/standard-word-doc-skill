#!/usr/bin/env python3
"""Render a formal DOCX from Markdown using the bundled standard template."""

from __future__ import annotations

import argparse
import copy
import re
import shutil
import uuid
import zipfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS, "r": R_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)

SKILL_DIR = Path(__file__).resolve().parents[1]
TEMPLATE = SKILL_DIR / "assets" / "standard-word-template.docx"
DEFAULT_OUTPUT_DIR = Path.home() / "Documents" / "AI-Stack-Outputs" / "word-docs"


@dataclass
class Block:
    kind: str
    text: str = ""
    level: int = 0
    rows: list[list[str]] | None = None


def qn(name: str) -> str:
    prefix, tag = name.split(":", 1)
    if prefix == "w":
        return f"{{{W_NS}}}{tag}"
    if prefix == "r":
        return f"{{{R_NS}}}{tag}"
    raise ValueError(name)


def parse_frontmatter(text: str) -> tuple[dict[str, str], str]:
    stripped = text.lstrip("\ufeff")
    if not stripped.startswith("---\n"):
        return {}, text
    end = stripped.find("\n---", 4)
    if end == -1:
        return {}, text
    raw = stripped[4:end].strip()
    body = stripped[stripped.find("\n", end + 1) + 1 :]
    meta: dict[str, str] = {}
    for line in raw.splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        meta[key.strip()] = value.strip().strip('"').strip("'")
    return meta, body


def split_table(lines: list[str], start: int) -> tuple[Block | None, int]:
    if start + 1 >= len(lines):
        return None, start
    if "|" not in lines[start] or "|" not in lines[start + 1]:
        return None, start
    separator = [c.strip() for c in lines[start + 1].strip().strip("|").split("|")]
    if not separator or not all(re.fullmatch(r":?-{3,}:?", c or "") for c in separator):
        return None, start
    rows: list[list[str]] = []
    i = start
    while i < len(lines) and "|" in lines[i]:
        rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
        i += 1
    if len(rows) >= 2:
        rows.pop(1)
    return Block(kind="table", rows=rows), i


def parse_markdown(text: str) -> tuple[dict[str, str], list[Block]]:
    meta, body = parse_frontmatter(text)
    lines = body.splitlines()
    blocks: list[Block] = []
    paragraph: list[str] = []

    def flush_paragraph() -> None:
        nonlocal paragraph
        if paragraph:
            blocks.append(Block(kind="paragraph", text=" ".join(paragraph).strip()))
            paragraph = []

    i = 0
    while i < len(lines):
        raw = lines[i].rstrip()
        line = raw.strip()
        if not line:
            flush_paragraph()
            i += 1
            continue
        table, new_i = split_table(lines, i)
        if table:
            flush_paragraph()
            blocks.append(table)
            i = new_i
            continue
        heading = re.match(r"^(#{1,5})\s+(.+)$", line)
        if heading:
            flush_paragraph()
            blocks.append(Block(kind="heading", level=len(heading.group(1)), text=heading.group(2).strip()))
            i += 1
            continue
        bullet = re.match(r"^[-*+]\s+(.+)$", line)
        if bullet:
            flush_paragraph()
            blocks.append(Block(kind="bullet", text=bullet.group(1).strip()))
            i += 1
            continue
        paragraph.append(line)
        i += 1
    flush_paragraph()
    return meta, blocks


def document_title(meta: dict[str, str], blocks: list[Block]) -> str:
    for key in ("title", "project_name", "项目名称"):
        if meta.get(key):
            return meta[key]
    for block in blocks:
        if block.kind == "heading" and block.level == 1:
            return block.text
    return "项目名称"


def document_type(meta: dict[str, str], blocks: list[Block]) -> str:
    for key in ("document_type", "doc_type", "文档类型"):
        if meta.get(key):
            return meta[key]
    for block in blocks:
        if block.kind == "heading" and block.level == 1:
            return block.text
    return "实施方案"


def sanitize_filename(text: str) -> str:
    text = re.sub(r"[\\/:*?\"<>|]+", "_", text.strip())
    text = re.sub(r"\s+", "_", text)
    return text[:80] or f"standard-doc-{uuid.uuid4().hex[:8]}"


def element_text(el: ET.Element) -> str:
    return "".join(t.text or "" for t in el.findall(".//w:t", NS))


def compact_text(text: str) -> str:
    return re.sub(r"\s+", "", text)


def replace_element_text(el: ET.Element, value: str) -> ET.Element:
    """Replace visible text while preserving the first run's formatting."""
    text_nodes = el.findall(".//w:t", NS)
    if text_nodes:
        text_nodes[0].text = value
        text_nodes[0].set(f"{{{XML_NS}}}space", "preserve")
        for node in text_nodes[1:]:
            node.text = ""
        return el

    first_run = el.find(".//w:r", NS)
    if first_run is None:
        if el.tag == qn("w:p"):
            first_run = ET.SubElement(el, qn("w:r"))
        elif el.find(".//w:p", NS) is not None:
            first_run = ET.SubElement(el.find(".//w:p", NS), qn("w:r"))  # type: ignore[arg-type]
        else:
            p = ET.SubElement(el, qn("w:p"))
            first_run = ET.SubElement(p, qn("w:r"))
    t = ET.SubElement(first_run, qn("w:t"), {f"{{{XML_NS}}}space": "preserve"})
    t.text = value
    return el


def clone_with_text(el: ET.Element, value: str) -> ET.Element:
    cloned = copy.deepcopy(el)
    strip_runtime_ids(cloned)
    replace_element_text(cloned, value)
    return cloned


def strip_runtime_ids(el: ET.Element) -> None:
    """Avoid duplicating Word/WPS paragraph identity attributes when cloning."""
    for node in el.iter():
        for attr in list(node.attrib):
            if attr.endswith("}paraId") or attr.endswith("}textId") or attr.endswith("}rsidR") or attr.endswith("}rsidRPr"):
                del node.attrib[attr]


def normalize_run_properties(rpr: ET.Element) -> None:
    fonts = rpr.find("w:rFonts", NS)
    if fonts is not None:
        fallback_font = (
            fonts.get(qn("w:eastAsia"))
            or fonts.get(qn("w:ascii"))
            or fonts.get(qn("w:hAnsi"))
            or fonts.get(qn("w:cs"))
        )
        if fallback_font:
            for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                fonts.set(qn(attr), fonts.get(qn(attr)) or fallback_font)

    size = rpr.find("w:sz", NS)
    if size is not None:
        size_value = size.get(qn("w:val"))
        complex_size = rpr.find("w:szCs", NS)
        if complex_size is None:
            ET.SubElement(rpr, qn("w:szCs"), {qn("w:val"): size_value or ""})
        elif size_value:
            complex_size.set(qn("w:val"), size_value)

    if rpr.find("w:b", NS) is not None and rpr.find("w:bCs", NS) is None:
        ET.SubElement(rpr, qn("w:bCs"))


def normalize_heading_paragraph(p: ET.Element) -> ET.Element:
    for rpr in p.findall(".//w:rPr", NS):
        normalize_run_properties(rpr)
    return p


def style_name_map(styles_xml: bytes | None) -> dict[str, str]:
    if not styles_xml:
        return {}
    try:
        root = ET.fromstring(styles_xml)
    except ET.ParseError:
        return {}
    styles: dict[str, str] = {}
    for style in root.findall("w:style", NS):
        style_id = style.get(qn("w:styleId"))
        name_el = style.find("w:name", NS)
        if style_id and name_el is not None:
            name = name_el.get(qn("w:val"))
            if name:
                styles[name.lower()] = style_id
    return styles


def style_id(styles: dict[str, str], names: tuple[str, ...], fallback: str) -> str:
    for name in names:
        found = styles.get(name.lower())
        if found:
            return found
    return fallback


def resolve_style_ids(styles_xml: bytes | None) -> dict[str, str]:
    styles = style_name_map(styles_xml)
    return {
        "heading1": style_id(styles, ("heading 1", "标题 1"), "2"),
        "heading2": style_id(styles, ("heading 2", "标题 2"), "3"),
        "heading3": style_id(styles, ("heading 3", "标题 3"), "4"),
        "heading4": style_id(styles, ("heading 4", "标题 4"), "5"),
        "body": style_id(styles, ("body ref", "正文", "normal"), "19"),
        "toc": style_id(styles, ("body ref", "toc 1", "目录 1"), "19"),
        "table": style_id(styles, ("table text ref", "表格正文"), "21"),
    }


def paragraph(text: str = "", style: str = "19", preserve_space: bool = True) -> ET.Element:
    p = ET.Element(qn("w:p"))
    ppr = ET.SubElement(p, qn("w:pPr"))
    ET.SubElement(ppr, qn("w:pStyle"), {qn("w:val"): style})
    if text:
        r = ET.SubElement(p, qn("w:r"))
        t_attrs = {f"{{{XML_NS}}}space": "preserve"} if preserve_space else {}
        t = ET.SubElement(r, qn("w:t"), t_attrs)
        t.text = text
    return p


def page_break() -> ET.Element:
    p = ET.Element(qn("w:p"))
    r = ET.SubElement(p, qn("w:r"))
    ET.SubElement(r, qn("w:br"), {qn("w:type"): "page"})
    return p


def heading(text: str, level: int, styles: dict[str, str], prototypes: dict[str, ET.Element]) -> ET.Element:
    style = styles.get(f"heading{min(max(level, 1), 4)}", "2")
    prototype = prototypes.get(style)
    if prototype is not None:
        return normalize_heading_paragraph(clone_with_text(prototype, text))
    return paragraph(text, style=style)


def bullet(text: str, styles: dict[str, str]) -> ET.Element:
    p = paragraph(text, style=styles.get("body", "19"))
    ppr = p.find("w:pPr", NS)
    assert ppr is not None
    numpr = ET.SubElement(ppr, qn("w:numPr"))
    ET.SubElement(numpr, qn("w:ilvl"), {qn("w:val"): "0"})
    ET.SubElement(numpr, qn("w:numId"), {qn("w:val"): "2"})
    return p


def toc_title_from_template(template_body: ET.Element, styles: dict[str, str]) -> ET.Element:
    for child in template_body:
        if child.tag == qn("w:p") and compact_text(element_text(child)) == "目录":
            return clone_with_text(child, "目  录")
    return paragraph("目  录", style=styles.get("toc", "19"))


def toc_paragraph(template_body: ET.Element, styles: dict[str, str]) -> list[ET.Element]:
    title = toc_title_from_template(template_body, styles)
    p = ET.Element(qn("w:p"))
    ppr = ET.SubElement(p, qn("w:pPr"))
    ET.SubElement(ppr, qn("w:pStyle"), {qn("w:val"): styles.get("toc", "19")})
    r1 = ET.SubElement(p, qn("w:r"))
    ET.SubElement(r1, qn("w:fldChar"), {qn("w:fldCharType"): "begin", qn("w:dirty"): "true"})
    r2 = ET.SubElement(p, qn("w:r"))
    instr = ET.SubElement(r2, qn("w:instrText"), {f"{{{XML_NS}}}space": "preserve"})
    instr.text = r'TOC \o "1-4" \h \z \u'
    r3 = ET.SubElement(p, qn("w:r"))
    ET.SubElement(r3, qn("w:fldChar"), {qn("w:fldCharType"): "separate"})
    r4 = ET.SubElement(p, qn("w:r"))
    t = ET.SubElement(r4, qn("w:t"))
    t.text = "目录将在最终生成时自动刷新"
    r5 = ET.SubElement(p, qn("w:r"))
    ET.SubElement(r5, qn("w:fldChar"), {qn("w:fldCharType"): "end"})
    return [title, p]


@dataclass
class TableTemplate:
    tbl_pr: ET.Element | None
    header_tc_pr: ET.Element | None
    body_tc_pr: ET.Element | None
    band_tc_pr: ET.Element | None
    header_p: ET.Element | None
    body_p: ET.Element | None
    band_p: ET.Element | None


def first_regular_table_template(template_body: ET.Element) -> TableTemplate:
    fallback = TableTemplate(None, None, None, None, None, None, None)
    for tbl in template_body.findall("w:tbl", NS):
        rows = tbl.findall("w:tr", NS)
        if len(rows) < 2:
            continue
        first_cells = rows[0].findall("w:tc", NS)
        second_cells = rows[1].findall("w:tc", NS)
        if len(first_cells) < 2 or len(second_cells) < 2:
            continue
        third_cells = rows[2].findall("w:tc", NS) if len(rows) > 2 else []
        return TableTemplate(
            tbl_pr=copy.deepcopy(tbl.find("w:tblPr", NS)),
            header_tc_pr=copy.deepcopy(first_cells[0].find("w:tcPr", NS)),
            body_tc_pr=copy.deepcopy(second_cells[0].find("w:tcPr", NS)),
            band_tc_pr=copy.deepcopy(third_cells[0].find("w:tcPr", NS)) if third_cells else None,
            header_p=copy.deepcopy(first_cells[0].find("w:p", NS)),
            body_p=copy.deepcopy(second_cells[0].find("w:p", NS)),
            band_p=copy.deepcopy(third_cells[0].find("w:p", NS)) if third_cells else None,
        )
    return fallback


def remove_child(parent: ET.Element, tag: str) -> None:
    for child in list(parent):
        if child.tag == qn(tag):
            parent.remove(child)


def apply_cell_width(tc_pr: ET.Element, width: int) -> None:
    remove_child(tc_pr, "w:tcW")
    ET.SubElement(tc_pr, qn("w:tcW"), {qn("w:w"): str(width), qn("w:type"): "dxa"})


def template_cell(text: str, tc_pr: ET.Element | None, p_template: ET.Element | None, style: str, width: int) -> ET.Element:
    tc = ET.Element(qn("w:tc"))
    if tc_pr is not None:
        tcpr = copy.deepcopy(tc_pr)
    else:
        tcpr = ET.Element(qn("w:tcPr"))
    apply_cell_width(tcpr, width)
    tc.append(tcpr)
    if p_template is not None:
        tc.append(clone_with_text(p_template, text))
    else:
        tc.append(paragraph(text, style=style))
    return tc


def table(rows: list[list[str]], table_template: TableTemplate, styles: dict[str, str]) -> ET.Element:
    tbl = ET.Element(qn("w:tbl"))
    if table_template.tbl_pr is not None:
        tbl.append(copy.deepcopy(table_template.tbl_pr))
    else:
        tblpr = ET.SubElement(tbl, qn("w:tblPr"))
        ET.SubElement(tblpr, qn("w:tblStyle"), {qn("w:val"): "11"})
        ET.SubElement(tblpr, qn("w:tblW"), {qn("w:w"): "9000", qn("w:type"): "dxa"})
        borders = ET.SubElement(tblpr, qn("w:tblBorders"))
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            ET.SubElement(borders, qn(f"w:{side}"), {qn("w:val"): "single", qn("w:sz"): "4", qn("w:space"): "0", qn("w:color"): "000000"})
    max_cols = max((len(row) for row in rows), default=1)
    width = 9000 // max_cols
    for row_index, row in enumerate(rows):
        tr = ET.SubElement(tbl, qn("w:tr"))
        for cell in row + [""] * (max_cols - len(row)):
            if row_index == 0:
                tc = template_cell(cell, table_template.header_tc_pr, table_template.header_p, styles.get("table", "21"), width)
            elif row_index % 2 == 0 and table_template.band_tc_pr is not None:
                tc = template_cell(cell, table_template.band_tc_pr, table_template.band_p, styles.get("table", "21"), width)
            else:
                tc = template_cell(cell, table_template.body_tc_pr, table_template.body_p, styles.get("table", "21"), width)
            tr.append(tc)
    return tbl


def section_properties_from_template(body: ET.Element) -> ET.Element:
    sect = body.find("w:sectPr", NS)
    if sect is None:
        sect = ET.Element(qn("w:sectPr"))
    return copy.deepcopy(sect)


def heading_prototypes_from_template(template_body: ET.Element, styles: dict[str, str]) -> dict[str, ET.Element]:
    heading_style_ids = {styles.get(f"heading{level}", "") for level in range(1, 5)}
    prototypes: dict[str, ET.Element] = {}
    for child in template_body:
        if child.tag != qn("w:p"):
            continue
        style_el = child.find("w:pPr/w:pStyle", NS)
        if style_el is None:
            continue
        style = style_el.get(qn("w:val"))
        if style in heading_style_ids and style not in prototypes and compact_text(element_text(child)):
            prototypes[style] = copy.deepcopy(child)
    return prototypes


def numbering_run_properties_by_style(numbering_xml: bytes | None) -> dict[str, ET.Element]:
    if not numbering_xml:
        return {}
    try:
        root = ET.fromstring(numbering_xml)
    except ET.ParseError:
        return {}

    result: dict[str, ET.Element] = {}
    for level in root.findall(".//w:lvl", NS):
        style_el = level.find("w:pStyle", NS)
        rpr = level.find("w:rPr", NS)
        if style_el is not None and rpr is not None:
            style = style_el.get(qn("w:val"))
            if style:
                result[style] = copy.deepcopy(rpr)
    return result


def ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(tag, NS)
    if child is None:
        child = ET.SubElement(parent, qn(tag))
    return child


def harmonize_bullet_numbering_definitions(numbering_xml: bytes | None) -> bytes | None:
    if not numbering_xml:
        return numbering_xml
    try:
        root = ET.fromstring(numbering_xml)
    except ET.ParseError:
        return numbering_xml

    for level in root.findall(".//w:lvl", NS):
        num_fmt = level.find("w:numFmt", NS)
        level_text = level.find("w:lvlText", NS)
        is_bullet = (
            num_fmt is not None
            and num_fmt.get(qn("w:val")) == "bullet"
            or level_text is not None
            and level_text.get(qn("w:val")) in {"●", "○", "•", "·"}
        )
        if not is_bullet:
            continue

        rpr = ensure_child(level, "w:rPr")
        fonts = ensure_child(rpr, "w:rFonts")
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            fonts.set(qn(attr), "微软雅黑")

        size = ensure_child(rpr, "w:sz")
        size.set(qn("w:val"), size.get(qn("w:val")) or "24")
        size_cs = ensure_child(rpr, "w:szCs")
        size_cs.set(qn("w:val"), size.get(qn("w:val")) or "24")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def harmonize_heading_style_definitions(
    styles_xml: bytes | None,
    numbering_xml: bytes | None,
    heading_styles: dict[str, str],
    prototypes: dict[str, ET.Element],
) -> bytes | None:
    if not styles_xml:
        return styles_xml
    try:
        root = ET.fromstring(styles_xml)
    except ET.ParseError:
        return styles_xml

    styles_by_id = {style.get(qn("w:styleId")): style for style in root.findall("w:style", NS)}
    numbering_rprs = numbering_run_properties_by_style(numbering_xml)
    style_ids = [heading_styles.get(f"heading{level}", "") for level in range(1, 5)]
    for style_id_value in style_ids:
        style = styles_by_id.get(style_id_value)
        prototype = prototypes.get(style_id_value)
        prototype_rpr = prototype.find("w:r/w:rPr", NS) if prototype is not None else None
        source_rpr = prototype_rpr if prototype_rpr is not None else numbering_rprs.get(style_id_value)
        if style is None or source_rpr is None:
            continue
        for child in list(style):
            if child.tag == qn("w:rPr"):
                style.remove(child)
        style_rpr = copy.deepcopy(source_rpr)
        normalize_run_properties(style_rpr)
        style.append(style_rpr)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def cover_elements_from_template(
    template_body: ET.Element,
    *,
    cover_title: str,
    doc_type: str,
    organization: str,
    classification: str,
    version: str,
    doc_date: str,
) -> list[ET.Element]:
    elements: list[ET.Element] = []
    found_toc = False
    for child in template_body:
        if child.tag == qn("w:p") and compact_text(element_text(child)) == "目录":
            found_toc = True
            break
        elements.append(copy.deepcopy(child))

    if not found_toc or not elements:
        return [
            paragraph(cover_title, style="10"),
            paragraph(doc_type, style="10"),
            paragraph(f"文件密级：{classification}", style="20"),
            paragraph(f"编制单位：{organization}", style="20"),
            paragraph(f"编制日期：{doc_date}", style="20"),
            paragraph(f"版本号：{version}", style="20"),
            page_break(),
        ]

    title_replaced = 0
    for el in elements:
        if el.tag != qn("w:p"):
            continue
        text = compact_text(element_text(el))
        if not text:
            continue
        if "文件密级" in text:
            replace_element_text(el, f"文件密级：{classification}")
        elif "编制单位" in text:
            replace_element_text(el, f"编制单位：{organization}")
        elif "编制日期" in text or text.startswith("日期"):
            replace_element_text(el, f"编制日期：{doc_date}")
        elif "版本号" in text or text.startswith("版本"):
            replace_element_text(el, f"版本号：{version}")
        elif title_replaced == 0:
            replace_element_text(el, cover_title)
            title_replaced += 1
        elif title_replaced == 1:
            replace_element_text(el, doc_type)
            title_replaced += 1
    return elements


def build_body(template_body: ET.Element, styles_xml: bytes | None, meta: dict[str, str], blocks: list[Block]) -> ET.Element:
    body = ET.Element(qn("w:body"))
    styles = resolve_style_ids(styles_xml)
    table_template = first_regular_table_template(template_body)
    heading_prototypes = heading_prototypes_from_template(template_body, styles)
    title = document_title(meta, blocks)
    doc_type = document_type(meta, blocks)
    cover_title = meta.get("cover_title") or meta.get("cover_project") or meta.get("封面标题") or title
    organization = meta.get("organization") or meta.get("编制单位") or "XXXX单位"
    classification = meta.get("classification") or meta.get("文件密级") or "内部资料"
    version = meta.get("version") or meta.get("版本") or "V1.0"
    doc_date = meta.get("date") or meta.get("日期") or f"{date.today():%Y年%m月%d日}"

    for item in cover_elements_from_template(
        template_body,
        cover_title=cover_title,
        doc_type=doc_type,
        organization=organization,
        classification=classification,
        version=version,
        doc_date=doc_date,
    ):
        body.append(item)
    for item in toc_paragraph(template_body, styles):
        body.append(item)
    body.append(page_break())
    body.append(paragraph("", style=styles.get("body", "19")))

    skipped_title = False
    for block in blocks:
        if block.kind == "heading":
            if block.level == 1 and not skipped_title and block.text == title:
                skipped_title = True
                continue
            body.append(heading(block.text, max(block.level - 1, 1), styles, heading_prototypes))
        elif block.kind == "paragraph":
            body.append(paragraph(block.text, style=styles.get("body", "19")))
        elif block.kind == "bullet":
            body.append(bullet(block.text, styles))
        elif block.kind == "table" and block.rows:
            body.append(table(block.rows, table_template, styles))

    body.append(section_properties_from_template(template_body))
    return body


def render(input_path: Path, output_dir: Path, filename: str | None, template_path: Path = TEMPLATE) -> Path:
    if not template_path.exists():
        raise SystemExit(f"template not found: {template_path}")
    meta, blocks = parse_markdown(input_path.read_text(encoding="utf-8"))
    if not blocks:
        raise SystemExit("input markdown has no content")

    output_dir.mkdir(parents=True, exist_ok=True)
    out_name = filename or f"{sanitize_filename(document_title(meta, blocks))}.docx"
    if not out_name.lower().endswith(".docx"):
        out_name += ".docx"
    output_path = output_dir / out_name

    shutil.copyfile(template_path, output_path)
    with zipfile.ZipFile(output_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    root = ET.fromstring(files["word/document.xml"])
    old_body = root.find("w:body", NS)
    assert old_body is not None
    template_body = copy.deepcopy(old_body)
    styles = resolve_style_ids(files.get("word/styles.xml"))
    heading_prototypes = heading_prototypes_from_template(template_body, styles)
    harmonized_styles = harmonize_heading_style_definitions(
        files.get("word/styles.xml"),
        files.get("word/numbering.xml"),
        styles,
        heading_prototypes,
    )
    if harmonized_styles is not None:
        files["word/styles.xml"] = harmonized_styles
    harmonized_numbering = harmonize_bullet_numbering_definitions(files.get("word/numbering.xml"))
    if harmonized_numbering is not None:
        files["word/numbering.xml"] = harmonized_numbering
    root.remove(old_body)
    root.append(build_body(template_body, files.get("word/styles.xml"), meta, blocks))
    files["word/document.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Markdown input")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--filename", help="Output file name")
    parser.add_argument("--template", type=Path, default=TEMPLATE, help="DOCX template to inherit")
    args = parser.parse_args()
    out = render(args.input, args.output_dir, args.filename, args.template)
    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
