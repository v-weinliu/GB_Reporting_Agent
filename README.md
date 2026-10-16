conda create -n py313 python=3.13
conda activate py313

pip install mcp pywin32 openpyxl

## Run and debug

Requires Microsoft Word on Windows and the VS Code Python Debugger extension.
Select the Python environment containing `pywin32` in VS Code.

Open Run and Debug (`Ctrl+Shift+D`), select **Merge Word reports**, and press
`F5`. Enter `input1`, `input2`, and `output` when prompted. Paths can be
absolute or relative to the workspace root; do not add surrounding quotes.
The first report is used as the base document, then the second is appended.
Existing output files are overwritten; the output must differ from both inputs.

Command-line equivalent:

```powershell
python .\.github\skills\merge-word-reports\merge_word_reports.py --input1 "first.docx" --input2 "second.docx" --output "merged.docx"
```

Omitting the arguments uses the original report filenames in the script directory.

## Merge skill

The [merge-word-reports skill](.github/skills/merge-word-reports/SKILL.md)
requires `input1` and `input2` to point to existing Word files. Its `output`
argument is optional: if no output filename is supplied, the skill generates
`merge-yyyyMMddHHmm.docx` using the actual local time at invocation.
The default output directory is the workspace root, unless the user specifies
another directory. This naming rule applies to the skill, not direct script
execution or the Run and Debug configuration.