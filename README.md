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
- Structured error handling via custom exceptions
- Filename-only outputs automatically placed next to the input file
- Safe low-level IPC layer between Python and native automation scripts

## Installation

### Using pip

```bash
pip install --upgrade conver
```

### Recommended for CLI use: pipx

If you primarily use `conver` as a command-line tool, it is best installed with **pipx**.  
This isolates the package in its own virtual environment and avoids polluting your system Python:

```bash
pipx install --upgrade conver
```

After installation:

```bash
conver --help
```

will be available globally.

---

### Requirements

- Python **3.9+**  
- Microsoft Word must be installed on the system  
- macOS (JXA) or Windows (PowerShell + Word COM automation)  

---

## Python Usage

### Basic example

```python
from conver import conver

conver("document.docx", "document.pdf")
```

Returns:

```python
Path("/absolute/path/document.pdf")
```

### Filename-only output

```python
conver("/Users/me/docs/a.docx", "a.pdf")
```

Automatically produces:

```bash
/Users/me/docs/a.pdf
```

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

#### Signature

conver(input_path, output_path, keep_open=False) → pathlib.Path

#### Parameters

- **input_path** — `str | Path`  
  Path to the source document.

- **output_path** — `str | Path`  
  Output filename or full path.

- **keep_open** — `bool`  
  Leave Microsoft Word running after conversion.

#### Returns

- **pathlib.Path** — absolute path of the generated output file.

#### Raises

- `InputFileNotFound`
- `UnsupportedFormat`
- `WordStartError`
- `SaveError`
- `IPCError`
- `PlatformNotSupported`

---

## CLI Usage

The CLI supports two modes: direct output path, or format flags.

### 1. Explicit input/output

```bash
conver input.docx output.pdf
```

### 2. Format flags (output placed next to input)

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
input.docx --pdf # -> input.pdf
input.docx --rtf # -> input.rtf
```

### Output Behavior

When OUTPUT is omitted, the resulting file is written next to the INPUT file.
Examples:

```bash
conver a.docx --pdf # -> a.pdf
conver /path/x.docx # -> /path/x.pdf
```

When multiple inputs are provided, --output must be a directory.  
If the directory does not exist, it will be created automatically.

### Default Format

If no format flag is provided and no explicit output path is given,
the default output format is **PDF**:

```bash
conver input.docx # -> input.pdf
```

### CLI Syntax

```
Usage:
    conver <input> <output>
    conver <input> [--pdf | --docx | --rtf | --txt | --html] [--keep-open]

Examples:
    conver a.docx a.pdf
    conver a.docx --pdf
    conver /path/to/file.doc --html
```

### CLI Arguments

`input`  
    One or more input files. Patterns like `*.docx` are expanded by the shell.

`output`  
    Optional.  
    For single input: output filename.  
    For multiple inputs: must be a directory.

---

### Multiple Inputs (Batch Mode)

The CLI can process multiple input files at once:

```bash
conver *.docx -o outdir
conver file1.docx file2.docx file3.docx -o converted/
```

Rules:

- When multiple inputs are provided, `--output` **must** point to a directory.
- If the directory does not exist, it will be created automatically.
- If no format flag is provided, the default output format is **PDF**.
- Format flags (`--pdf`, `--rtf`, etc.) **cannot be used together with** `--output FILE`.
- Globbing (`*.docx`) is expanded by your shell before reaching the CLI.

Output filenames follow the pattern:

```
input.docx  ->  outdir/input.pdf
```

### Shell Globbing

Patterns like `*.docx` are expanded by your shell before the `conver` command is executed.

Examples:
```bash
conver *.docx -o outdir
conver path/*.rtf --pdf
```

This means the CLI receives the expanded list of files as separate arguments.

### Format-Flag Restrictions

Format-selection flags: `--pdf`, `--docx`, `--rtf`, `--txt`, `--html`

Work **only when OUTPUT is omitted**, i.e.:

*Allowed*:  
```bash
conver input.docx --pdf
conver input.docx output.pdf
```

*Invalid*:  
```bash
conver input.docx output.pdf --pdf
```

---

## Supported Formats

Conversion capabilities depend on Microsoft Word.

Input formats: `.docx`, `.doc`, `.pdf`, `.rtf`, `.odt`, `.txt`, `.html`

Output formats: `.pdf`, `.docx`, `.doc`, `.pdf`, `.rtf`, `.odt`, `.txt`, `.html`

---

## Platform Notes

### macOS (JXA)

The script `convert.jxa` runs through:

```bash
osascript -l JavaScript
```

macOS may require granting Microsoft Word file-access permissions.

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
