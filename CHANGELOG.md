# Conver Changelog

## [0.1.1] - 2024-11-17

### Added
- CLI: automatic output directory inference when `--output` is omitted:
  - If all input files share the same parent directory, outputs are written there.
  - If inputs originate from different directories, the CLI now requires `--output DIRECTORY`.

### Improved
- Updated README with new batch-mode output behavior.
- Added clearer error message for ambiguous multi-directory inputs.

## [0.1.0] - 2024-11-16

### Added
- Initial public release. [7fd7291]
- PyPI packaging and script entry point.
- CLI tool `conver`:
  - Single-file conversion
  - Format flags (`--pdf`, `--rtf`, `--txt`, `--html`, `--docx`)
  - Batch mode with multiple inputs and `--output DIR`
- Structured error model (`ConverError` + specific subclasses).
- High-level Python API: `conver()` for Word document conversion.
- Reliable IPC layer with normalized JSON protocol.
- Native automation scripts:
  - macOS: `convert.jxa` (JXA / osascript)
  - Windows: `convert.ps1` (PowerShell + Word COM)
