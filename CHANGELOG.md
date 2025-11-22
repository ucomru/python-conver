# Conver Changelog

## [0.1.2] - 2025-11-22

### Added
- CLI: added long-form output format flags `--doc` and `--odt` for DOC/ODT conversion
  (short flags not assigned to avoid conflicts). [3d3b623]

### Improved
- Batch mode optimized: Word remains open during multi-file conversion,
  reducing startup overhead (final file honors user `--keep-open` flag). [a292633]
- CLI: output/format handling: [dc43909]
    - Refined output-path logic for both single-file and batch modes:
      unified validation of file vs directory semantics, correct directory auto-creation,
      and consistent suffix-based detection of explicit file paths.
      Added robust separation of cases where --output is a file or a directory.
    - Removed obsolete conflicting checks between --output and format flags,
      replacing them with correct single-mode validation rules.
    - Made format-flag resolution consistent across all branches
      (pdf default fallback preserved).
    - Added support for --doc and --odt output formats.
- Unified behavior of native conversion scripts: removed obsolete macOS-specific logic
  and updated USAGE blocks for consistent, readable cross-platform operation. [878ffb2]

## [0.1.1] - 2025-11-17

### Added
- CLI: automatic output directory inference when `--output` is omitted:
  - If all input files share the same parent directory, outputs are written there.
  - If inputs originate from different directories, the CLI now requires `--output DIRECTORY`.

### Improved
- Updated README with new batch-mode output behavior.
- Added clearer error message for ambiguous multi-directory inputs.

## [0.1.0] - 2025-11-16

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
