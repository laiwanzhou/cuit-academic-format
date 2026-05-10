# cuit-academic-format

用于成都信息工程大学论文格式检查与修复的 skill 仓库。

## 当前 Skill

- `cuit-thesis-docx-format`
  - 入口文件：`cuit-thesis-docx-format/SKILL.md`
  - 核心脚本：`cuit-thesis-docx-format/scripts/cuit_thesis_docx_format.py`

## 说明

该 skill 会对 `.docx` 论文执行格式检查，并输出：

- `*_format_comments.docx`
- `*_format_fixed.docx`
- `*_format_report.json`
- `*_format_report.html`

已包含最新修复：

- 目录 TOC 域识别（避免误报“缺失目录”）
- 摘要中英文切换不再误报模块重复
- 页码分段边界修复（前置部分与正文分离）
