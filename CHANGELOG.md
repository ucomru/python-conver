# Conver Changelog

## [0.1.6] - 2026-11-07

### Fixed
- CLI: fixed Microsoft Word temporary file filtering by correcting extension matching
  (missing leading dots). [a2fe1f5]
- Windows: fixed DOC input format validation by checking for null instead
  of treating format code `0` as invalid. [3aa808a]
- macOS: improved Word cleanup on conversion errors while preserving efficient
  batch conversion and automatic application shutdown. [52e427e]
- Windows: ensured Word documents and application instances are closed
  on conversion errors when `keepOpen` is disabled. [03ab659]
- README: corrected CLI examples to use `-o` / `--output`
  and updated the documented command syntax. [a708b9d]

## [0.1.5] - 2026-11-07

### Added
- CLI: added `-q` / `--quiet` option to suppress successful conversion output
  while preserving error messages. [b123440]
- Added PEP 561 `py.typed` marker to distribute inline type annotations
  for static type checkers. [df2e11a]

## [0.1.4] - 2026-11-06

### Added
- CLI: automatic skipping of Microsoft Word temporary files
  (`~$*.docx`, `~$*.doc`, `~$*.rtf`, `~$*.txt`, `~$*.html`, `~$*.odt`).
- Added `--include-temp` flag to disable temporary file filtering.

### Fixed
- Path normalization: filenames starting with `~` are no longer incorrectly interpreted
  as home-directory paths. Home expansion is now limited to `~` and `~/...`.

### Changed
- Raised minimum supported Python version from 3.9 to 3.12.

### Improved
- Updated type annotations across CLI, API, and IPC layers for stricter static type checking.
- Refined subprocess handling and JSON response typing.
- Updated README and cleaned up module documentation.

## [0.1.3] - 2025-11-22

### Improved
- CLI: updated short-flag layout for all formats; resolved conflicts with `-h`, `-o`, and `-d`.
  Introduced new short flags: `-x` (DOCX), `-d` (DOC), `-m` (HTML), `-u` (ODT).

### Changed
- CLI: restored `-h` as the standard help flag; switched version flag to `-v`.
  Added `help_option_names` for consistent `-h/--help` behavior.

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
