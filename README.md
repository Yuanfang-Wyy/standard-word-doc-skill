# Standard Word Doc Skill

用于 Codex 的企业级 Word 文档格式生成 Skill。它的核心目标不是自由设计版式，而是基于脱敏 Word 模板稳定生成正式 `.docx` 材料。

## 安装

推荐方式：

```bash
cd /path/to/ai-stack-control
./skill add https://github.com/Yuanfang-Wyy/standard-word-doc-skill
./skill sync
```

手动方式：

```bash
mkdir -p ~/.codex/skills
git clone https://github.com/Yuanfang-Wyy/standard-word-doc-skill.git ~/.codex/skills/standard-word-doc
```

## 使用

在 Codex 中直接说：

```text
[$standard-word-doc] 根据这些要点生成一份正式实施方案 Word
```

或运行脚本：

```bash
python3 scripts/render_standard_docx.py examples/sanitized-template-source.md
```

默认输出目录：

```text
~/Documents/AI-Stack-Outputs/word-docs
```

## 说明

- `assets/standard-word-template.docx` 是脱敏模板，只保留格式、标题体系、自动编号、表格样式、页眉页脚和目录字段。
- 默认不依赖 LibreOffice、Word 或 WPS 自动化。
- 如果安装 LibreOffice，可使用 `scripts/finalize_docx.sh` 尝试自动刷新目录和字段。
