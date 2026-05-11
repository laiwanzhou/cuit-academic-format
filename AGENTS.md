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