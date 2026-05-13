---
name: standard-word-doc
description: Use when creating, auditing, or standardizing formal Chinese Word materials such as implementation plans, necessity reports, construction plans, test reports, service summaries, and official review documents that should follow the standard Word template.
---

# Standard Word Doc

## Purpose

Use this skill to produce or inspect formal `.docx` materials that must follow the standard Word template in `assets/standard-word-template.docx`.

The generator is a document template engine, not a visual design engine. Format fidelity has higher priority than content richness: all generated documents must inherit the bundled template's cover, heading styles, numbering, body style, table style, page settings, header/footer, TOC field and page-number fields.

Default generation has no LibreOffice dependency. It creates a standard `.docx` from the bundled template and tells the user whether the TOC/page fields may need updating in Word/WPS. If LibreOffice is available, `scripts/finalize_docx.sh` can optionally refresh fields automatically.

## Common Requests

- "按标准规范写一份 Word"
- "把这个材料调整成标准规范格式"
- "根据这些要点生成正式实施方案"
- "检查这份 Word 是否符合标准模板"
- "把 Markdown 草稿转成正式报审 Word"

## Inputs

For generation, prefer Markdown with:

- One `#` title for the project or document title
- Optional frontmatter or key-value metadata for `document_type`, `subtitle`, `organization`, `classification`, `version`, `date`
- `##` to `#####` headings for chapter levels
- Paragraphs, bullet lists, and simple Markdown tables

For auditing, provide a `.docx` file path.

## Workflow

1. Read `references/style-guide.md` for formatting rules.
2. If document type matters, read `references/document-types.md`.
3. Run `scripts/check_dependencies.sh` to verify Python and the bundled template.
4. For generation, run `scripts/render_standard_docx.py` to create a `.docx`; do not hand-design cover, headings or tables outside the template.
5. If LibreOffice is available and the user wants automatic field refresh, run `scripts/finalize_docx.sh`.
6. Deliver the generated file path and clearly state whether field refresh was automatic or should be done in Word/WPS.
7. For auditing, run `scripts/audit_standard_docx.py` and return the audit report path plus the highest-risk findings.

## Required Commands

Generate:

```bash
python3 scripts/render_standard_docx.py input.md --output-dir "$HOME/Documents/AI-Stack-Outputs/word-docs"
```

Audit:

```bash
python3 scripts/audit_standard_docx.py input.docx --output "$HOME/Documents/AI-Stack-Outputs/word-docs/audit.md"
```

Dependency check:

```bash
scripts/check_dependencies.sh
```

## Hard Rules

- Format priority is higher than content priority. If there is a conflict, preserve the Word template.
- The bundled template is the only visual standard. Do not create new title, table, cover, header/footer or color styles in the generator.
- Use Heading 1 to Heading 4 from the template for document hierarchy. Do not hand-write chapter numbers into Markdown.
- Generate headings by cloning the template's heading paragraph prototypes, replacing text, and synchronizing heading style definitions with the prototype or numbering run properties so automatic numbering and heading text keep the same font, size and paragraph rules in Word/WPS.
- Tables must inherit the template's table properties, first-row cell properties, body cell properties and alternating row properties when available.
- Bullet numbering must not use Symbol or Wingdings fonts, because WPS may render those bullets as garbled characters. Normalize bullet numbering fonts during generation.
- Body copy must be formal, concise, professional, consulting-style and technical-document-style; avoid 口语化、自媒体语气、ChatGPT 式总结腔 and repetitive "首先、其次、最后".
- Do not use Pages as the finalization engine.
- Do not require LibreOffice for default generation.
- If LibreOffice finalization was not run, explicitly say the document was generated from the standard template and the TOC/page fields can be updated in Word/WPS if needed.
- Do not overwrite the user's source files. Write outputs to `$HOME/Documents/AI-Stack-Outputs/word-docs` unless the user explicitly requests another path.
- If the user requests automatic TOC/page refresh and LibreOffice is missing, explain the optional install command: `brew install --cask libreoffice`.
