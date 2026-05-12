# Codex 工作记忆：中文编码与低风险修改要求

本仓库包含大量中文规范、中文报告文案、中文测试样例。所有修改必须严格保持 UTF-8 编码，禁止产生中文乱码、问号串或 mojibake。

## 一、编码硬性要求

1. 所有源码、JSON、Markdown、测试文件必须按 UTF-8 读取和写入。
2. 修改 Python 文件时，必须使用：
   - `Path.read_text(encoding="utf-8")`
   - `Path.write_text(text, encoding="utf-8", newline="\n")`
3. 修改 JSON 文件时，必须使用：
   - `json.loads(path.read_text(encoding="utf-8"))`
   - `json.dumps(data, ensure_ascii=False, indent=2)`
   - `path.write_text(..., encoding="utf-8", newline="\n")`
4. 禁止使用系统默认编码读取或写入中文文件。
5. 禁止使用会破坏中文的 shell 重定向、批量替换、未知编码转换工具。
6. 禁止把正常中文替换成：
   - `????`
   - `??`
   - `浣滆`
   - `鏀昏`
   - `??????`
   - `\ufffd`
7. 除非是测试 Unicode 转义本身，禁止把面向用户的中文报告文案写成不可读的 Unicode escape。报告、JSON expected、HTML、Word 批注文案必须是正常中文。

## 二、Windows PowerShell 执行前设置

在执行含中文路径、中文输出、中文测试数据的命令前，先运行：

```powershell
$env:PYTHONIOENCODING='utf-8'
$OutputEncoding=[System.Text.Encoding]::UTF8
[Console]::OutputEncoding=[System.Text.Encoding]::UTF8
```

不要依赖 Windows 默认代码页。

## 三、修改方式要求

1. 不要进行全文件批量替换。
2. 不要大范围重写 `scripts/cuit_thesis_docx_format.py`。
3. 不要全局替换历史乱码字符串。
4. 只在与当前任务直接相关的函数附近做小块修改。
5. 每完成一个小块修改，必须立即运行：

```powershell
python -m py_compile cuit-thesis-docx-format/scripts/cuit_thesis_docx_format.py
```

6. 如果 `py_compile` 失败，立刻停止，先修复语法错误，不要继续叠加修改。
7. 如果测试失败，输出失败原因，不要 commit，不要 push。

## 四、提交前必须检查中文是否损坏

提交前必须搜索这些危险模式：

```powershell
Select-String -Path cuit-thesis-docx-format\**\* -Pattern "????","??????","浣滆","鏀昏"," " -SimpleMatch
```

如果发现这些内容出现在源码、规则、Markdown、测试样例、expected 文案、报告文案中，必须修复后再提交。

特别检查：

```powershell
python -m json.tool cuit-thesis-docx-format/references/rules.json > $null
python -m py_compile cuit-thesis-docx-format/scripts/cuit_thesis_docx_format.py
cd cuit-thesis-docx-format
python -m pytest tests
```

## 五、rules.json 中文文案要求

`cuit-thesis-docx-format/references/rules.json` 中的 `expected` 必须是可读中文，不能出现问号串。

以下文案必须保持正常中文。

### body.expected

```text
正文：中文宋体小四12pt，英文和数字Times New Roman 12pt，两端对齐，首行缩进2个汉字符，段前段后0磅，固定20磅行距。
```

### figure_caption.expected

```text
图序、图名、图注：宋体五号10.5pt，居中，固定20磅行距，段前0行，段后1行，置于图下方，图序与图名之间空一格。
```

### table_caption.expected

```text
表序、表名、表注：宋体五号10.5pt，居中，固定20磅行距，段前1行，段后0行，置于表上方，表序与表名之间空一格。
```

## 六、参考文献检查中文关键字

参考文献检查逻辑中，禁止使用 `??`、`[????]` 这类占位字符串判断中文语义。

应使用正常中文或稳定正则，例如：

```text
引用日期
学校
大学
出版年
出版地
访问路径
获取和访问路径
数字对象唯一标识符
DOI
等
```

参考文献测试样例也必须使用正常中文，例如：

```text
[1] 作者. 题名[J]. 期刊名, 2020, 12(3): 15-20.
[2] 作者. 题名[EB/OL]. (2020-01-01)[2024-01-01]. https://example.com.
[3] 作者. 题名[D]. 成都: 成都信息工程大学, 2023.
```

不要使用：

```text
[1] ??. ??. https://example.com
[????]
??
```

## 七、九大模块边界要求

论文只允许九个连续模块：

```text
cover -> declaration -> abstract -> toc -> symbols -> body -> references -> appendix -> acknowledgement
```

`author_bio` 不得作为第十个独立模块。

如果检测到：

```text
作者简历
攻读学位期间发表
学术论文与研究成果
```

且不在 appendix 模块内，应报告：

```text
检测到九大模块之外的疑似额外组成部分，需要人工确认。
```

不要把这类内容切换成独立模块。

## 八、自动修改边界

以下内容可以自动修改：

1. 正文 body paragraph 两端对齐。
2. 正文 body paragraph 中汉字两端多余空格。
3. 摘要正文、正文、致谢正文中的安全空段。
4. 图片高度超过 16cm 时，按比例缩放到 16cm。

以下内容原则上只报告，不自动大改：

1. 参考文献学术内容。
2. 图片不是 6cm × 8cm，但高度不超过 16cm。
3. 表格是否三线表。
4. 表格续表和表头重复情况。
5. 任何需要人工判断的语义规则。

## 九、提交前最终命令

提交前必须运行：

```powershell
cd D:\work\codex_agent\cuit-academic-format-clean

git status --short

python -m json.tool cuit-thesis-docx-format/references/rules.json > $null
python -m py_compile cuit-thesis-docx-format/scripts/cuit_thesis_docx_format.py

cd cuit-thesis-docx-format
python -m pytest tests
```

如果任何命令失败，不要 commit，不要 push。

## 十、禁止提交临时文件

不要提交：

```text
__pycache__/
.pytest_cache/
tmp-enhanced-format-test/
tmp-align-test/
tmp-section-boundary-test/
*.tmp
临时 DOCX
临时 HTML
临时 JSON 报告
```

提交前如有这些文件，先清理或停止询问。

## 十一、当前仓库已知问题

最新提交后，仓库中可能已经存在中文损坏问题，需要优先修复。

重点检查：

1. `cuit-thesis-docx-format/references/rules.json`
   - `body.expected`
   - `figure_caption.expected`
   - `table_caption.expected`

   这些字段不能是 `????` 或问号串，必须恢复为正常中文。

2. `cuit-thesis-docx-format/scripts/cuit_thesis_docx_format.py`
   - 参考文献检查逻辑不能使用 `??`、`[????]` 判断。
   - 表格、续表、引用日期、学校、大学、出版地等中文语义判断必须是正常中文或可靠正则。

3. `cuit-thesis-docx-format/tests/test_enhanced_rules.py`
   - 测试样例不能使用 `??` 作为中文占位。
   - 应使用真实中文参考文献样例。

修复中文乱码和问号占位是优先级最高的任务。
