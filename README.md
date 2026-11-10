# conver — Word document converter

[![CI](https://github.com/ucomru/python-conver/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/ucomru/python-conver/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/conver)](https://pypi.org/project/conver/)
[![Python](https://img.shields.io/badge/python-%3E%3D3.12-blue)](https://pypi.org/project/conver/)
[![License](https://img.shields.io/github/license/ucomru/python-conver)](https://github.com/ucomru/python-conver/blob/main/LICENSE)


`conver` is a cross-platform Python package that converts Microsoft Word documents  
using native system automation:

- macOS: JXA (JavaScript for Automation)
- Windows: PowerShell + Word COM automation

The package provides:

1. A high-level Python API: `conver()` and `conver_stage()`
2. A command-line interface: `conver`
3. A low-level IPC layer calling platform scripts  
   (`convert.jxa` for macOS, `convert.ps1` for Windows)

Conversion is performed via Microsoft Word itself, ensuring maximum compatibility 
with `.docx`, `.doc`, `.rtf`, `.txt`, `.html`, `.htm`, `.odt`.

---

## Features

- Cross-platform: macOS and Windows
- Uses the actual Microsoft Word engine
- Clean, minimal Python API
- CLI with direct input/output or format-selection flags
- Optional quiet mode for scripting and batch conversion
- Optional filtering of input files by supported extensions
- Optional macOS staging for both Python API and CLI
  to reduce Microsoft Word output-directory access prompts
- Structured error handling via custom exceptions
- Filename-only outputs automatically placed next to the input file
- Safe low-level IPC layer between Python and native automation scripts

---

## Installation

### Using pip

```bash
pip install --upgrade conver
```

### Recommended for CLI use: pipx

If you primarily use `conver` as a command-line tool,
we recommend installing it with [pipx](https://pipx.pypa.io/).

It installs applications in isolated virtual environments without affecting
your system Python.

*Install:*

```bash
pipx install conver
```

*Upgrade:*

```bash
pipx upgrade conver
```

After installation, the `conver` command will be available globally:

```bash
conver --help
```

If `pipx` is *not installed*, see the
[official installation guide](https://pipx.pypa.io/stable/how-to/install-pipx.html).

### Requirements

- Python **3.12+**  
- Microsoft Word must be installed on the system  
- macOS (JXA) or Windows (PowerShell + Word COM automation)  

---

## Python Usage

### Basic example

```python
from conver import conver

conver("document.docx", "document.pdf")
```

Returns: `Path("/absolute/path/document.pdf")`

### Convert DOCX to PDF with pathlib

Use `Path.with_suffix()` to create the PDF next to the source document:

```python
from pathlib import Path
from conver import conver

source = Path("documents/report.docx")
result = conver(source, source.with_suffix(".pdf"))
```

The resulting PDF is saved as `documents/report.pdf`.

### Filename-only output

```python
conver("/Users/me/docs/document.docx", "document.pdf")
```

Automatically produces: `/Users/me/docs/document.pdf`

### Staging conversion (macOS only)

On macOS, Microsoft Word may request permission to save files in certain directories.
The `conver_stage()` function converts the document in a temporary directory before
moving the result to its final destination.

```python
from conver import conver_stage

conver_stage("document.docx", "document.pdf")
```

By default, the staging directory is `~/Downloads`.

To specify a different directory:

```python
conver_stage("document.docx", "document.pdf", stage_dir="~/Documents/WordTemp")
```

The staging directory must already exist and be accessible to Microsoft Word.

Staging is not supported for HTML output (`.html`, `.htm`).  
Use `conver()` for these formats.

Staging-related failures raise `StagingError`:

```python
from conver import conver_stage, StagingError

try:
    conver_stage("document.docx", "document.pdf")
except StagingError as err:
    print(f"Staging failed: {err}")
```

### Error handling example

```python
from conver import conver, UnsupportedFormat

try:
    conver("document.doc", "document.xyz")
except UnsupportedFormat:
    print("This output format is not supported.")
```

---

## Python API Reference

### `conver()` function

High-level document conversion API.

*Signature:*

```python
conver(input=input_path, output=output_path, keep_open=False) -> pathlib.Path
```

*Parameters:*

- `input: str | Path`  
  Path to the source document.
- `output: str | Path`  
  Output filename or full path.
- `keep_open: bool`  
  Leave Microsoft Word running after conversion.

*Returns:*

- `pathlib.Path`  
  Absolute path of the generated output file.

*Raises:*

- `InputFileNotFound`
- `UnsupportedFormat`
- `WordStartError`
- `SaveError`
- `IPCError`
- `PlatformNotSupported`

### `conver_stage()` function

High-level staged document conversion API (macOS only).

Converts a document in a temporary directory, then moves the result to
the requested output path. This can reduce Microsoft Word file-access
permission prompts.

*Signature:*

```python
conver_stage(input, output, keep_open=False, stage_dir=None) -> pathlib.Path
```

*Parameters:*

- `input: str | Path`
  Path to the source document.
- `output: str | Path`
  Output filename or full path.
- `keep_open: bool`
  Leave Microsoft Word running after conversion. Default: `False`.
- `stage_dir: str | Path | None`
  Temporary output directory. Default: `~/Downloads`.

*Returns:*

- `pathlib.Path`
  Absolute path of the generated output file.

*Raises:*

- `InputFileNotFound`
- `UnsupportedFormat`
- `WordStartError`
- `SaveError`
- `IPCError`
- `PlatformNotSupported`
- `StagingError`

*Notes:*

- Supported only on macOS.
- HTML output (`.html`, `.htm`) is not supported.
- The staging directory must already exist.
- Existing output files are replaced only after conversion and transfer.
- Overwriting the source document is prohibited.
- Temporary staging files are removed on normal completion and ordinary failures.

---

## CLI Usage

The CLI supports two modes: direct output path, or format flags.

### Explicit input/output

```bash
conver document.docx -o file_name.pdf
```

### Format flags (output placed next to input)

```bash
conver document.docx --pdf # -> document.pdf
conver document.docx --doc # -> document.doc
conver document.docx --rtf # -> document.rtf
conver document.docx --txt # -> document.txt
conver document.docx --html # -> document.html
conver document.docx --odt # -> document.odt
```

### Output Behavior

When OUTPUT is omitted, the resulting file is written next to the INPUT file.

*Examples:*

```bash
conver document.docx # -> document.pdf
conver document.docx --pdf # -> document.pdf
conver /path/document.docx # -> /path/document.pdf
```

When multiple inputs are provided, `-o`/`--output`, if specified, must be a directory.
If the directory does not exist, it will be created automatically.

### Default Format

If no format flag is provided and no explicit output path is given,
the default output format is **PDF**:

```bash
conver document.docx # -> document.pdf
```

### Quiet Mode

By default, the CLI prints the output path after each successful conversion.

Use `-q` or `--quiet` to suppress success output. Error messages
are still displayed, and failures return a non-zero exit code.

```bash
conver document.docx --pdf --quiet
conver *.docx -q
```

### Supported Extensions Only

Use `-e` or `--ext-only` to process only files with supported input extensions.

This is useful when converting documents from directories containing other
file types or subdirectories.

```bash
conver * -e
conver * --ext-only
```

Only regular files with supported extensions
(`.docx`, `.doc`, `.rtf`, `.txt`, `.html`, `.htm`, `.odt`)
are processed. Extension matching is case-insensitive.

Microsoft Word temporary files (`~$...`) are still excluded by default
unless `--include-temp` is specified.

### Staging Mode (macOS only)

Microsoft Word for macOS may request permission to save files in certain directories.

Use `-s` or `--stage` to save converted files to a temporary location before
moving them to the final output directory.

```bash
conver document.docx --pdf -s
conver *.docx --pdf -o converted/ --stage
```

The default staging directory is `~/Downloads`.
Use `--stage-dir` to specify a different directory:

```bash
conver document.docx --pdf -s --stage-dir ~/Documents/WordTemp
```

Staging can also be configured through environment variables:

```bash
export CONVER_STAGE=1
export CONVER_STAGE_DIR="$HOME/Downloads"
```

Use `--no-stage` to explicitly disable staging, including when enabled
through `CONVER_STAGE`.

**Notes:**
- Staging is supported only on macOS.
- The staging directory must already exist and be accessible to Microsoft Word.
- In the CLI, HTML output (`.html` and `.htm`) bypasses staging and
  uses direct conversion.
- In the Python API, `conver_stage()` raises `UnsupportedFormat` for HTML output.
- Temporary output files are deleted after conversion or failure.
- Staging reduces output-directory permission prompts but does not suppress
  all Microsoft Word dialogs.

### CLI Syntax

*Usage:*

```
conver INPUT... [-o OUTPUT] \
    [--pdf | --docx | --doc | --rtf | --txt | --html | --odt] \
    [-k] [-q] [-e] [-s] [--stage-dir PATH]
```

*Examples:*

```bash
conver document.docx -o file_name.pdf
conver document.docx --pdf
conver /path/to/document.doc --html
conver *.docx -o converted/
```

### CLI Arguments

`INPUT`  
    One or more input files. Patterns like `*.docx` are expanded by the shell.

`-o [--output] OUTPUT`  
    Optional.  
    For a single input: output filename or directory.  
    For multiple inputs: must be a directory.

`-e, --ext-only`  
    Process only regular files with supported input extensions.  
    Skip unsupported file types and directories.

`-s, --stage`
    Enable macOS staging mode. Can also be enabled using `CONVER_STAGE=1`.

`--no-stage`
    Explicitly disable staging, overriding `CONVER_STAGE`.

`--stage-dir PATH`
    Temporary output directory. Default: `~/Downloads`.
    Can also be set using `CONVER_STAGE_DIR`.
    Does not enable staging by itself.

### Multiple Inputs (Batch Mode)

The CLI can process multiple input files at once:

```bash
conver *.docx -o outdir
conver file1.docx file2.docx file3.docx -o converted/
```

*Rules:*

- When multiple inputs are provided, `-o`/`--output`, if specified,
  must point to a directory.
- If the directory does not exist, it will be created automatically.
- If no format flag is provided, the default output format is **PDF**.
- Format flags (`--pdf`, `--rtf`, etc.) **cannot be used together with** `--output FILE`.
- Globbing (`*.docx`) is expanded by your shell before reaching the CLI.

If multiple input files are provided and `-o`/`--output` is **not** specified,
the CLI automatically attempts to infer a common parent directory for all inputs.
If all files reside in the same directory, that directory becomes the output location.

**Warning:** In batch mode, output filenames are derived from input filenames.
Files with the same stem but different extensions (`document.docx` and `document.rtf`)
produce the same output filename (`document.pdf`).
Later conversions may overwrite earlier results.

Use broad shell patterns such as `*` with care. The `-e, --ext-only` option filters
unsupported input types but does not prevent output filename collisions.

*Example:*

```bash
conver *.docx
# -> writes output next to each input file
```

If the input files come from **different** directories, the output directory
becomes ambiguous and the CLI will require an explicit `-o`/`--output`.

### Shell Globbing

Patterns like `*.docx` are expanded by your shell before the `conver`
command is executed.

Examples:
```bash
conver *.docx -o output_dir
conver path/*.rtf --pdf
```

This means the CLI receives the expanded list of files as separate arguments.

Use `conver * -e` to process only supported document types in a mixed directory.

Note that shell wildcards such as `*` typically exclude hidden files.

By default, the CLI skips Microsoft Word temporary files
(`~$*.docx`, `~$*.doc`, `~$*.rtf`, `~$*.txt`, `~$*.html`, `~$*.htm`, `~$*.odt`).

Use `--include-temp` to disable this filter.

### Format-Flag Restrictions

Format-selection flags (`--pdf`, `--docx`, `--doc`, `--rtf`, `--txt`, `--html`, `--odt`)
can be combined with an output directory, but not with an explicit output file path.

*Allowed:*

```bash
conver document.doc
conver document.docx --pdf
conver document.docx -o file_name.pdf
conver *.docx -o converted/ --pdf
```

*Invalid:*

```bash
conver document.docx -o file_name.pdf --pdf
```

---

## Supported Formats

Conversion capabilities depend on Microsoft Word.

Input formats: `.docx`, `.doc`, `.rtf`, `.txt`, `.html`, `.htm`, `.odt`

Output formats: `.pdf`, `.docx`, `.doc`, `.rtf`, `.txt`, `.html`, `.htm`, `.odt`

---

## Platform Notes

### macOS (JXA)

The script `convert.jxa` runs through:

```bash
osascript -l JavaScript
```

Microsoft Word may request file-access permissions for output directories.
The optional staging mode can reduce these prompts.
See [Staging Mode](#staging-mode-macos-only).

### Windows (PowerShell)

The script `convert.ps1` uses:

```powershell
Word.Application COM automation
```

If Word prompts the user, automation must be allowed.

---

## Low-Level Error Codes (Reference)

These codes come from platform scripts and are mapped to exceptions by `conver()`:

| Code | Meaning                        |
|------|--------------------------------|
| 0    | Success                        |
| 1    | Invalid JSON from stdin        |
| 2    | Unsupported input format       |
| 3    | Unsupported output format      |
| 11   | Input file not found           |
| 21   | Word startup timeout           |
| 31   | Word could not save the file   |
| 98   | Script produced invalid JSON   |
| 99   | Unsupported platform           |

`StagingError` is raised directly by the Python API and has no low-level IPC error code.

---

## Development and Testing

Install the project with test dependencies:

```bash
pip install -e ".[test]"
```

Run the test suite:

```bash
pytest -q
```

The tests cover the Python API, CLI, staging, and IPC handling. They use mocks
and temporary files, so Microsoft Word is not required to run the test suite.

Real document conversion requires Microsoft Word on macOS or Windows.

---

## License

MIT License  
(c) 2024 Timur Ulyahin  
https://github.com/ucomru

---

## Project Homepage

GitHub:  
https://github.com/ucomru/python-conver

PyPI:  
https://pypi.org/project/conver/
