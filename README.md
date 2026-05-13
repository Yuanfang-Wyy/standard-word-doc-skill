# Standard Word Doc Skill

一个通用的企业级 Word 文档格式生成 Skill。它面向支持 `SKILL.md` 结构的 AI Agent / Skill Runtime，不绑定某一个具体工具。

这个 Skill 的核心目标不是“自由生成内容”，而是基于脱敏 Word 模板稳定生成正式 `.docx` 材料，让输出结果更接近企业内部正式模板中的人工文档。

## 适用场景

- 生成实施方案、建设方案、必要性分析、测试报告、总结报告、服务单材料等正式 Word 文档。
- 将 Markdown 草稿转换成符合标准模板的 `.docx`。
- 审计已有 `.docx` 是否包含封面字段、标题层级、目录字段、编号体系和页脚页码。
- 团队统一 Word 材料格式，减少人工排版和模板走样。

## 安装

### 方式一：通过命令行安装

如果你的环境支持标准 Skills CLI，可直接通过 GitHub URL 安装：

```bash
npx skills add https://github.com/Yuanfang-Wyy/standard-word-doc-skill
```

安装完成后，重启或刷新你的 AI 工具，使其重新加载 Skills。

### 方式二：让 AI Agent 代为安装

在支持 Skills 的 AI 工具中，可以直接对 Agent 说：

```text
请安装这个 Skill：https://github.com/Yuanfang-Wyy/standard-word-doc-skill
```

Agent 应保持仓库目录结构完整，包括：

```text
SKILL.md
assets/
references/
scripts/
examples/
```

### 方式三：手动安装

如果你的工具暂不支持上述安装方式，可将本仓库完整下载或 clone 到该工具指定的 Skills 目录中。不同工具的目录不同，请以对应工具文档为准。

```bash
git clone https://github.com/Yuanfang-Wyy/standard-word-doc-skill.git
```

常见做法是将整个仓库放入你的 AI 工具的 skills 目录，并确保工具能读取仓库根目录下的 `SKILL.md`。

### 可选：通过 SkillHub 安装

如果你的团队已经使用 SkillHub，也可以通过 SkillHub 搜索和安装该技能。SkillHub 不是必需依赖。

## 使用

在任意支持 Skills 的 AI 工具中，可以这样描述任务：

```text
使用 standard-word-doc，根据这些要点生成一份正式实施方案 Word。
```

也可以直接运行脚本生成示例文档：

```bash
python3 scripts/render_standard_docx.py examples/sanitized-template-source.md
```

默认输出目录：

```text
~/Documents/AI-Stack-Outputs/word-docs
```

## 输入格式

推荐使用 Markdown 作为中间结构层：

```markdown
---
title: 项目名称
cover_title: 项目名称
document_type: 实施方案
organization: 编制单位
classification: 内部资料
version: V1.0
date: 二〇二六年
---

# 项目名称

## 项目背景

正文内容。

### 建设现状

正文内容。
```

标题层级会映射到模板中的 `Heading 1` 到 `Heading 4`，自动编号由 Word 多级编号体系产生，不建议在 Markdown 中手写“第一章”“1.1”等编号。

## 模板说明

- `assets/standard-word-template.docx` 是脱敏模板，只保留通用版式、标题体系、自动编号、表格样式、页眉页脚和目录字段。
- 公开仓库不包含项目正文、客户名称、个人路径、服务器地址、Token 或内部材料。
- 如需替换为团队自己的模板，请保持文件名为 `assets/standard-word-template.docx`，并确保模板中存在标题样式、正文样式、表格样式、目录字段和页脚页码。

## 依赖说明

默认只需要：

```bash
python3
```

检查依赖：

```bash
scripts/check_dependencies.sh
```

默认生成流程不依赖 LibreOffice、Microsoft Word 或 WPS 自动化。若安装 LibreOffice，可选用以下脚本尝试自动刷新目录和字段：

```bash
scripts/finalize_docx.sh input.docx final.docx
```

## 目录结构

```text
standard-word-doc-skill/
├── SKILL.md
├── assets/
│   └── standard-word-template.docx
├── examples/
│   └── sanitized-template-source.md
├── references/
│   ├── document-types.md
│   └── style-guide.md
└── scripts/
    ├── audit_standard_docx.py
    ├── check_dependencies.sh
    ├── finalize_docx.sh
    └── render_standard_docx.py
```
