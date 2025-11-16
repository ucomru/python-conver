"""
Convenient public API for the conver package.
"""

from .conver import conver
from .conver import ConverError
from .conver import InputFileNotFound
from .conver import UnsupportedFormat
from .conver import WordStartError
from .conver import SaveError
from .conver import IPCError
from .conver import PlatformNotSupported
from .__version__ import __version__

__all__ = [
    "conver",
    "ConverError",
    "InputFileNotFound",
    "UnsupportedFormat",
    "WordStartError",
    "SaveError",
    "IPCError",
    "PlatformNotSupported",
    "__version__",
]
