---
name: merge-word-reports
description: '当用户要求合并两个 Word 报告、合并 Word 文件或调用 merge_word_reports.py 时使用。收集 input1、input2 和可选的 output，未指定输出文件名时按当前时间生成 merge-yyyyMMddHHmm.docx，验证两个输入文件真实存在，再调用本 Skill 目录下的脚本生成合并后的 Word 文件。'
argument-hint: 'input1=<第一个 Word 文件路径> input2=<第二个 Word 文件路径> [output=<输出 Word 文件路径>]'
user-invocable: true
---

# 合并 Word 报告

## 目的

调用本 Skill 目录下、与本文档同级的 [merge_word_reports.py](./merge_word_reports.py)，
以第一个 Word 报告为基础，追加第二个 Word 报告，并按脚本现有逻辑处理标题编号和测试结果，生成输出文件。
不要重新实现合并逻辑，也不要复制或修改脚本。

## 参数

| 参数 | 含义 | 要求 |
| --- | --- | --- |
| `input1` | 第一个 Word 文件的路径 | 必须是真实存在的 Word 文件，作为合并后的基础文档 |
| `input2` | 第二个 Word 文件的路径 | 必须是真实存在的 Word 文件，内容追加到第一个文档后 |
| `output` | 生成的 Word 文件的路径，可选 | 指定文件名时以 `.docx` 结尾；未指定文件名时默认使用 `merge-yyyyMMddHHmm.docx`，不能与任一输入文件相同 |

## 执行流程

### 1. 提醒用户指定路径

- 调用时提醒用户：“请指定 input1（第一个 Word 文件）和 input2（第二个 Word 文件）的路径；output 可选，未指定输出文件名时会根据当前时间自动生成。”
- 如果缺少 `input1` 或 `input2`，使用可用的结构化提问工具一次性收集缺失的输入路径；两个输入路径未齐全前不得运行脚本。
- 不得猜测输入文件或使用脚本内置的默认输入路径。
- 未提供 `output` 时，在项目根目录生成默认输出文件；如果用户只指定输出目录，则在该目录下生成默认文件名，无需再询问文件名。
- 默认文件名格式为 `merge-yyyyMMddHHmm.docx`，即四位年、两位月、两位日、24 小时制两位小时、两位分钟连续拼接，不加分隔符。
  执行前通过系统时钟读取实际本地时间，例如 PowerShell 使用 `"merge-$(Get-Date -Format 'yyyyMMddHHmm').docx"`；
  不得使用示例时间、历史消息时间或模型推测的时间。每次调用只生成一次文件名，后续检查、执行和反馈使用同一个输出路径。
- 用户指定完整输出文件名时保留该名称，不要替换为默认名称。
- 确定三个最终路径后，复述输入路径和输出路径并继续检查；仅缺少 `output` 时无需提问。
- 支持绝对路径和相对项目根目录的路径；执行前统一解析为绝对路径。
- 如果用户提供的路径有歧义，先澄清，不要自行选择文件。

### 2. 检查文件和输出位置

必须先检查，再调用脚本：

1. 确认本 Skill 目录下存在 `merge_word_reports.py`，在当前项目中即 `.github\skills\merge-word-reports\merge_word_reports.py`。
2. 分别检查 `input1` 和 `input2`，确认路径存在且指向普通文件，而不是目录。
   在 PowerShell 中使用 `Test-Path -LiteralPath <路径> -PathType Leaf`，不要只检查扩展名。
3. 确认输入文件扩展名为 `.doc` 或 `.docx`，不区分大小写。文件存在和扩展名正确并不保证内容有效；能否由 Word 打开以脚本实际执行结果为准。
4. 任一输入文件不存在、不是文件或类型不符时，明确指出参数名和失败路径，请用户提供正确路径，然后重新检查。检查失败时不得运行脚本。
5. 确认 `output` 以 `.docx` 结尾，且其父目录存在。父目录不存在时请用户更正，或经用户明确同意后创建。
6. 比较解析后的绝对路径，按 Windows 路径大小写不敏感规则，禁止 `output` 与任一输入文件相同。
   如果输出文件已经存在，还要检查是否通过符号链接、硬链接等方式指向输入文件，避免覆盖源文档。
7. 如果 `output` 已存在且是目录，要求用户更正；如果是文件，提醒用户脚本会删除并覆盖该文件。
   只有用户明确同意覆盖后才能继续；否则请用户指定新路径。不要自行删除文件。

### 3. 检查运行环境

- 本脚本依赖 Windows、已安装的 Microsoft Word 和 `pywin32`。
- 使用当前项目已选择的 Python 解释器；不要假设终端中的 `python` 一定指向该环境。
- 如果运行时报告缺少依赖或 Word 不可用，明确报告问题，不得声称合并成功。
- 不得未经同意切换环境、安装软件或结束用户的 Word 进程。

### 4. 调用脚本

在项目根目录执行，将三个路径作为独立参数传入：

```powershell
python .\.github\skills\merge-word-reports\merge_word_reports.py --input1 "D:\Reports\first.docx" --input2 "D:\Reports\second.docx" --output "D:\Reports\merged.docx"
```

以上路径仅为示例，必须替换为用户指定或按上述规则生成并通过检查的绝对路径；
即使用户省略 `output`，也必须将生成的输出路径显式传给 `--output`，不得使用脚本内置的输出默认值。
`python` 也应替换为当前项目实际使用的解释器命令。
带空格或特殊字符的路径必须作为单个参数传递，不要将用户输入拼接成可执行的命令字符串。
使用 PowerShell 时，优先通过调用运算符 `&` 和参数数组调用解释器。

### 5. 验证并反馈结果

- 等待脚本执行完成，检查退出码，并确认最终确定的 `output` 真实存在、是文件且大小大于零。
- 仅当退出码为零且输出检查通过时，报告生成成功，并提供输出文件的完整路径链接。
- 执行失败时，报告实际错误以及相关路径；即使输出文件已经出现，也不能将可能未完成的文件视为合并成功。
- 不要自行清理输入文件或失败时产生的输出文件。

## 调用示例

```text
/merge-word-reports input1="D:\Reports\PowerPoint.docx" input2="D:\Reports\Word.docx" output="D:\Reports\combined.docx"
```

省略输出路径时：

```text
/merge-word-reports input1="D:\Reports\PowerPoint.docx" input2="D:\Reports\Word.docx"
```

在项目根目录使用实际运行时间生成 `merge-yyyyMMddHHmm.docx`。
如果用户只输入 `/merge-word-reports`，先收集两个输入路径，自动确定默认输出路径，完成检查后再执行。
