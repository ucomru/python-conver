"""Public API for document conversion using Microsoft Word."""

from .__version__ import __version__
from .conver import (
    ConverError,
    InputFileNotFound,
    IPCError,
    PlatformNotSupported,
    SaveError,
    StagingError,
    UnsupportedFormat,
    WordStartError,
    conver,
    conver_stage,
)

__all__ = [
    "ConverError",
    "IPCError",
    "InputFileNotFound",
    "PlatformNotSupported",
    "SaveError",
    "StagingError",
    "UnsupportedFormat",
    "WordStartError",
    "__version__",
    "conver",
    "conver_stage",
]
