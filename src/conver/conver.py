# Author: Timur Ulyahin, https://github.com/ucomru
# License: MIT – provided "as-is" without any warranty or liability
# Copyright: (c) 2024 Timur Ulyahin

"""
High-level document conversion API.

This module exposes the `conver()` function — a Pythonic wrapper around the
low-level IPC-based converter. It normalizes paths, performs minimal validation,
invokes the underlying platform-specific script, and raises structured exceptions
on failure.

Successful conversion returns the resolved output path as a `Path` object.
Errors are represented as subclasses of `ConverError`.
"""

from pathlib import Path

from ._convert import convert


class ConverError(Exception):
    error_code: int | None

    def __init__(self, message: str, error_code: int | None = None) -> None:
        super().__init__(message)
        self.error_code = error_code


class InputFileNotFound(ConverError):
    pass


class UnsupportedFormat(ConverError):
    pass


class WordStartError(ConverError):
    pass


class SaveError(ConverError):
    pass


class IPCError(ConverError):
    pass


class PlatformNotSupported(ConverError):
    pass


_ERROR_MAP = {
    1: IPCError,
    2: UnsupportedFormat,
    3: UnsupportedFormat,
    11: InputFileNotFound,
    21: WordStartError,
    31: SaveError,
    98: IPCError,
    99: PlatformNotSupported,
}


def _expanduser(path: str | Path) -> Path:
    path = Path(path)
    return path.expanduser() if path.parts and path.parts[0] == "~" else path


def _normalize_paths(
    input: str | Path,
    output: str | Path,
) -> tuple[Path, Path]:
    """
    Normalize and resolve input/output paths.

    Rules:
    - input is converted to absolute Path.
    - output:
        * if only a filename is provided, it is placed into the same directory
          as input;
        * otherwise, it is resolved relative to the filesystem.
    - All returned paths are absolute.
    """

    # Normalize input
    input_path = _expanduser(input).absolute()

    # Normalize output raw
    output_raw = _expanduser(output)

    # filename-only -> drop into input directory
    if output_raw.parent == Path("."):
        output_path = input_path.parent / output_raw
    else:
        output_path = output_raw

    output_path = output_path.absolute()

    return input_path, output_path


def conver(
    input: str | Path,
    output: str | Path,
    keep_open: bool = False,
) -> Path:
    """
    Convert a document to another format via the platform-level converter.

    Parameters
    ----------
    input : str or pathlib.Path
        Path to the source document.
    output : str or pathlib.Path
        Target output path with desired extension.
    keep_open : bool, default=False
        Whether to keep Microsoft Word open after conversion.

    Returns
    -------
    Path
        The resolved absolute output path.

    Raises
    ------
    InputFileNotFound
        The input file does not exist.
    UnsupportedFormat
        The source or target format is not supported by the underlying converter.
    WordStartError
        Microsoft Word could not be started by the script.
    SaveError
        The document could not be saved in the target format.
    IPCError
        Communication with the script failed or produced invalid JSON.
    PlatformNotSupported
        The current OS is not supported.
    """

    input_path, output_path = _normalize_paths(input, output)

    if not input_path.exists():
        raise InputFileNotFound(f"Input file does not exist: {input_path}", 11)

    result = convert(
        input=str(input_path),
        output=str(output_path),
        keep_open=keep_open,
    )

    code = result.get("error_code")
    msg = result.get("message") or "Unknown error"

    if code != 0 or result.get("status") != "success":
        code = code or 98
        exc_class = _ERROR_MAP.get(code, IPCError)
        raise exc_class(f"[{code}] {msg}", code)

    return output_path
