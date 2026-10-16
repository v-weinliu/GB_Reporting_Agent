---
name: merge-word-reports
description: 'Use when the user asks to merge two Word reports or Word files, or to run merge_word_reports.py. Collect input1, input2, and an optional output. If no output filename is specified, generate merge-yyyyMMddHHmm.docx using the current time. Verify that both input files exist, then run the script in this skill directory to create the merged Word file.'
argument-hint: 'input1=<first Word file path> input2=<second Word file path> [output=<output Word file path>]'
user-invocable: true
---

# Merge Word Reports

## Purpose

Run [merge_word_reports.py](./merge_word_reports.py), located alongside this document in the skill directory.
Use the first Word report as the base, append the second report, and generate the output using the script's existing heading-numbering and test-result logic.
Do not reimplement the merge logic or copy or modify the script.

## Parameters

| Parameter | Meaning | Requirements |
| --- | --- | --- |
| `input1` | Path to the first Word file | Required; must be an existing Word file, used as the base document |
| `input2` | Path to the second Word file | Required; must be an existing Word file, appended to the first document |
| `output` | Path to the generated Word file, optional | An explicit filename must end in `.docx`; if no filename is provided, use `merge-yyyyMMddHHmm.docx`. Must not refer to either input file |

## Workflow

### 1. Collect the paths

- Remind the user: "Please specify input1 (the first Word file) and input2 (the second Word file). The output is optional; if no output filename is specified, one will be generated using the current time."
- If `input1` or `input2` is missing, use an available structured question tool to collect all missing input paths at once. Do not run the script until both input paths are available.
- Do not guess input files or use the script's built-in default input paths.
- If `output` is omitted, place the default output file in the project root. If the user specifies only an output directory, generate the default filename in that directory without asking for a filename.
- The default filename is `merge-yyyyMMddHHmm.docx`: a four-digit year, two-digit month, two-digit day, two-digit hour in 24-hour format, and two-digit minute, concatenated without separators.
  Read the actual local time from the system clock before execution, for example with `"merge-$(Get-Date -Format 'yyyyMMddHHmm').docx"` in PowerShell.
  Do not use example timestamps, historical message timestamps, or model-estimated times. Generate the filename once per invocation and use the same output path for validation, execution, and reporting.
- Preserve an explicit output filename supplied by the user instead of replacing it with the default.
- Once all three final paths are determined, restate the input and output paths and proceed with validation. Do not ask a question solely because `output` is missing.
- Accept absolute paths and paths relative to the project root; resolve all paths to absolute paths before execution.
- If a supplied path is ambiguous, ask for clarification rather than choosing a file.

### 2. Validate the files and output location

Complete these checks before running the script:

1. Confirm that `merge_word_reports.py` exists in this skill directory: `.github\skills\merge-word-reports\merge_word_reports.py` in the current project.
2. Check `input1` and `input2` separately to confirm that each path exists and refers to a file, not a directory.
   In PowerShell, use `Test-Path -LiteralPath <path> -PathType Leaf`; checking the extension alone is insufficient.
3. Confirm that each input has a `.doc` or `.docx` extension, ignoring case. Existence and a valid extension do not guarantee valid contents; whether Word can open the file is determined by actual script execution.
4. If either input is missing, is not a file, or has an unsupported extension, identify the parameter and failing path, request a corrected path, and validate again. Do not run the script if validation fails.
5. Confirm that `output` ends in `.docx` and its parent directory exists. If the parent directory is missing, request a corrected path or create the directory only with the user's explicit approval.
6. Compare the resolved absolute paths using case-insensitive Windows path rules. Do not allow `output` to refer to either input file.
   If the output file already exists, also check whether it refers to an input through a symbolic link, hard link, or another alias to avoid overwriting a source document.
7. If `output` already exists as a directory, request a corrected path. If it is a file, warn that the script will delete and overwrite it.
   Proceed only with the user's explicit approval to overwrite; otherwise request a new path. Do not delete files yourself.

### 3. Check the runtime environment

- The script requires Windows, an installed copy of Microsoft Word, and `pywin32`.
- Use the Python interpreter selected for the current project. Do not assume that `python` in the terminal points to that environment.
- If execution reports a missing dependency or unavailable Word installation, report the issue explicitly and do not claim that the merge succeeded.
- Do not switch environments, install software, or terminate the user's Word processes without permission.

### 4. Run the script

Run from the project root, passing the three paths as separate arguments:

```powershell
python .\.github\skills\merge-word-reports\merge_word_reports.py --input1 "D:\Reports\first.docx" --input2 "D:\Reports\second.docx" --output "D:\Reports\merged.docx"
```

The paths above are examples only. Replace them with validated absolute paths supplied by the user or generated according to the rules above.
Even if the user omits `output`, explicitly pass the generated output path to `--output`; do not use the script's built-in default output.
Replace `python` with the interpreter command actually used by the current project.
Pass paths containing spaces or special characters as single arguments. Do not concatenate user input into an executable command string.
In PowerShell, prefer invoking the interpreter with the call operator `&` and an argument array.

### 5. Verify and report the result

- Wait for the script to finish, check the exit code, and confirm that the final `output` exists, is a file, and has a size greater than zero.
- Report success only if the exit code is zero and the output checks pass. Provide a link to the output file using its full path.
- If execution fails, report the actual error and relevant paths. An output file may be incomplete even if it exists; do not treat its presence alone as a successful merge.
- Do not delete input files or output files left behind by a failed run.

## Invocation examples

```text
/merge-word-reports input1="D:\Reports\PowerPoint.docx" input2="D:\Reports\Word.docx" output="D:\Reports\combined.docx"
```

With the output path omitted:

```text
/merge-word-reports input1="D:\Reports\PowerPoint.docx" input2="D:\Reports\Word.docx"
```

Generate `merge-yyyyMMddHHmm.docx` in the project root using the actual execution time.
If the user enters only `/merge-word-reports`, collect both input paths, determine the default output path automatically, and complete validation before execution.
