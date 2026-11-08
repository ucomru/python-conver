# conver — Word document converter

`conver` is a cross-platform Python package that converts Microsoft Word documents  
using native system automation:

- macOS: JXA (JavaScript for Automation)
- Windows: PowerShell + Word COM automation

The package provides:

1. A high-level Python API: `conver.conver()`
2. A command-line interface: `conver`
3. A low-level IPC layer calling platform scripts  
   (`convert.jxa` for macOS, `convert.ps1` for Windows)

Conversion is performed via Microsoft Word itself, ensuring maximum compatibility 
with `.docx`, `.doc`, `.rtf`, `.txt`, `.html`, `.odt`.

---

## Features

- Cross-platform: macOS and Windows
- Uses the actual Microsoft Word engine
- Clean, minimal Python API
- CLI with direct input/output or format-selection flags
- Optional quiet mode for scripting and batch conversion
- Optional staging mode on macOS to reduce Microsoft Word output-directory access prompts
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

If you primarily use `conver` as a command-line tool, we recommend installing it
with [pipx](https://pipx.pypa.io/).  

It installs applications in isolated virtual environments without affecting your system Python.

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

If `pipx` is *not installed*,
see the [official installation guide](https://pipx.pypa.io/stable/how-to/install-pipx.html).

### Requirements

- Python **3.12+**  
- Microsoft Word must be installed on the system  
- MacOS (JXA) or Windows (PowerShell + Word COM automation)  

---

## Python Usage

### Basic example

```python
from conver import conver

conver("document.docx", "document.pdf")
```

Returns: `Path("/absolute/path/document.pdf")`

### Filename-only output

```python
conver("/Users/me/docs/a.docx", "a.pdf")
```

Automatically produces: `/Users/me/docs/a.pdf`

### Error handling example

```python
from conver import conver, UnsupportedFormat

try:
    conver("a.doc", "a.xyz")
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

---

## CLI Usage

The CLI supports two modes: direct output path, or format flags.

### Explicit input/output

```bash
conver input.docx -o output.pdf
```

### Format flags (output placed next to input)

```bash
conver input.docx --pdf
conver input.docx --rtf
conver input.docx --txt
conver input.docx --html
conver input.docx --docx
```

If only one input file and no explicit output is given,
the selected format determines the extension:

```bash
conver input.docx --pdf # -> input.pdf
conver input.docx --rtf # -> input.rtf
```

### Output Behavior

When OUTPUT is omitted, the resulting file is written next to the INPUT file.

*Examples:*

```bash
conver a.docx --pdf # -> a.pdf
conver /path/x.docx # -> /path/x.pdf
```

When multiple inputs are provided, `--output`, if specified, must be a directory.
If the directory does not exist, it will be created automatically.

### Default Format

If no format flag is provided and no explicit output path is given,
the default output format is **PDF**:

```bash
conver input.docx # -> input.pdf
```

### Quiet Mode

By default, the CLI prints the output path after each successful conversion.

Use `-q` or `--quiet` to suppress success output. Error messages
are still displayed, and failures return a non-zero exit code.

```bash
conver input.docx --pdf --quiet
conver *.docx -q
```

### Staging Mode (macOS only)

Microsoft Word for macOS may request permission to save files in certain directories.

Use `-s` or `--stage` to save converted files to a temporary location before moving them
to the final output directory.

```bash
conver input.docx --pdf -s
conver *.docx --pdf -o converted/ --stage
```

The default staging directory is `~/Downloads`.
Use `--stage-dir` to specify a different directory:

```bash
conver input.docx --pdf -s --stage-dir ~/Documents/WordTemp
```

Staging can also be configured through environment variables:

```bash
export CONVER_STAGE=1
export CONVER_STAGE_DIR="$HOME/Downloads"
```

Use `--no-stage` to explicitly disable staging, including when enabled through `CONVER_STAGE`.

**Notes:**
- Staging is supported only on macOS.
- The staging directory must already exist and be accessible to Microsoft Word.
- HTML output (`.html` and `.htm`) always uses direct conversion; staging is bypassed.
- Temporary output files are deleted after conversion or failure.
- Staging reduces output-directory permission prompts but does not suppress
  all Microsoft Word dialogs.

### CLI Syntax

*Usage:*

```
conver INPUT... [-o OUTPUT] \
    [--pdf | --docx | --doc | --rtf | --txt | --html | --odt] \
    [-k] [-q] [-s] [--stage-dir PATH]
```

*Examples:*

```bash
conver a.docx -o a.pdf
conver a.docx --pdf
conver /path/to/file.doc --html
conver *.docx -o converted/
```

### CLI Arguments

`INPUT`  
    One or more input files. Patterns like `*.docx` are expanded by the shell.

`-o [--output] OUTPUT`  
    Optional.  
    For single input: output filename.  
    For multiple inputs: must be a directory.

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

- When multiple inputs are provided, `--output`, if specified, must point to a directory.
- If the directory does not exist, it will be created automatically.
- If no format flag is provided, the default output format is **PDF**.
- Format flags (`--pdf`, `--rtf`, etc.) **cannot be used together with** `--output FILE`.
- Globbing (`*.docx`) is expanded by your shell before reaching the CLI.

If multiple input files are provided and `--output` is **not** specified,
the CLI automatically attempts to infer a common parent directory for all inputs.
If all files reside in the same directory, that directory becomes the output location.

*Example:*

```bash
conver *.docx
# -> writes output next to each input file
```

If the input files come from **different** directories, the output directory becomes ambiguous
and the CLI will require an explicit `--output`.

### Shell Globbing

Patterns like `*.docx` are expanded by your shell before the `conver` command is executed.

Examples:
```bash
conver *.docx -o outdir
conver path/*.rtf --pdf
```

This means the CLI receives the expanded list of files as separate arguments.

By default, the CLI skips Microsoft Word temporary files
(`~$*.docx`, `~$*.doc`, `~$*.rtf`, `~$*.txt`, `~$*.html`, `~$*.htm`, `~$*.odt`).

Use `--include-temp` to disable this filter.

### Format-Flag Restrictions

Format-selection flags (`--pdf`, `--docx`, `--doc`, `--rtf`, `--txt`, `--html`, `--odt`)
can be combined with an output directory, but not with an explicit output file path.

*Allowed:*

```bash
conver input.docx --pdf
conver input.docx -o output.pdf
conver *.docx -o converted/ --pdf
```

*Invalid:*

```bash
conver input.docx -o output.pdf --pdf
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
