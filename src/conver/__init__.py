"""Public API for document conversion using Microsoft Word."""

from .__version__ import __version__
from .conver import (
    ConverError,
    InputFileNotFound,
    IPCError,
    PlatformNotSupported,
    SaveError,
    UnsupportedFormat,
    WordStartError,
    conver,
)

__all__ = [
    "ConverError",
    "IPCError",
    "InputFileNotFound",
    "PlatformNotSupported",
    "SaveError",
    "UnsupportedFormat",
    "WordStartError",
    "__version__",
    "conver",
]
