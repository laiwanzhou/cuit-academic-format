# cuit-academic-format

## cuit-thesis-docx-format

用于检查和修复成都信息工程大学本科论文 `.docx` 格式的 skill。

- Skill 文件：`cuit-thesis-docx-format/SKILL.md`
- 核心脚本：`cuit-thesis-docx-format/scripts/cuit_thesis_docx_format.py`
- 规则文件：`cuit-thesis-docx-format/references/rules.json`

## 适用范围

- 仅支持 `.docx` 论文
- 不用于 `.pdf`、`.tex`、`.typ`

## 输出契约

每次执行固定生成：

- `<stem>_format_comments.docx`
- `<stem>_format_fixed.docx`
- `<stem>_format_report.json`
- `<stem>_format_report.html`

不会修改原始输入文件。

## 快速运行

```powershell
python cuit-thesis-docx-format/scripts/cuit_thesis_docx_format.py thesis.docx --output-dir ./res --renderer auto --screenshots auto
```

批量运行：

```powershell
python cuit-thesis-docx-format/scripts/run_batch_matrix.py "a.docx" "b.docx" --output-dir "./resx" --renderer ooxml --screenshots auto
```

## 环境检查

先执行：

```powershell
python cuit-thesis-docx-format/scripts/check_environment.py
```

能力级别：

- `full`：依赖完整，可使用 Office/WPS COM
- `ooxml-only`：仅 OOXML 能力
- `blocked`：依赖缺失

## 渲染与风险

- `--renderer auto|office|wps|ooxml`
- `--screenshots auto|require|never`
- `--allow-ooxml-layout-fixes` 仅在接受分页漂移风险时启用

OOXML-only 模式下，页码/页眉页脚/分节等高风险布局修复默认不自动改写，仅报告。

## 论文结构规则（核心）

论文按九个连续模块顺序识别：

1. 封面
2. 封二（声明）
3. 摘要与关键词
4. 目录
5. 符号说明（可选）
6. 正文
7. 参考文献
8. 附录（可选）
9. 致谢

规则按模块生效，后续模块不会复用前一模块规则。

页码规则：

- 声明-目录：小写罗马数字 `i, ii, iii...`
- 正文-致谢：`第*页 共*页`（阿拉伯数字）

## 当前版本关键修复

- 目录 TOC 域识别接入结构分析，避免误报“缺失目录”
- 中文/英文摘要切换不再误报“重复进入摘要模块”
- 正文与前置部分分节边界修复，避免页码格式串段

## 参考

详细行为、限制与完整工作流请查看：

- `cuit-thesis-docx-format/SKILL.md`
- `cuit-thesis-docx-format/tests/pressure-scenarios.md`
