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

from typing import Union, Tuple
from pathlib import Path

from ._convert import convert


class ConverError(Exception):
    def __init__(self, message, error_code=None):
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


def _normalize_paths(
    input_path: Union[str, Path],
    output_path: Union[str, Path],
) -> Tuple[Path, Path]:
    """
    Normalize and resolve input/output paths.

    Rules:
    - input_path is converted to absolute Path.
    - output_path:
        * if only a filename is provided, it is placed into the same directory
          as input_path;
        * otherwise, it is resolved relative to the filesystem.
    - All returned paths are absolute.
    """

    # Normalize input
    in_path = Path(input_path).expanduser().absolute()

    # Normalize output raw
    out_raw = Path(output_path).expanduser()

    # filename-only -> drop into input directory
    if out_raw.parent == Path("."):
        out_path = in_path.parent / out_raw
    else:
        out_path = out_raw

    out_path = out_path.absolute()

    return in_path, out_path


def conver(
    input_path: Union[str, Path],
    output_path: Union[str, Path],
    keep_open: bool = False,
) -> Path:
    """
    Convert a document to another format via the platform-level converter.

    Parameters
    ----------
    input_path : str or pathlib.Path
        Path to the source document.
    output_path : str or pathlib.Path
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

    in_path, out_path = _normalize_paths(input_path, output_path)

    if not in_path.exists():
        raise InputFileNotFound(f"Input file does not exist: {in_path}", 11)

    result = convert(
        input_path=str(in_path),
        output_path=str(out_path),
        keep_open=keep_open,
    )

    code = result.get("error_code", None)
    msg = result.get("message", "Unknown error")

    if code and code != 0:
        exc_class = _ERROR_MAP.get(code, ConverError)
        raise exc_class(f"[{code}] {msg}", code)

    return out_path
