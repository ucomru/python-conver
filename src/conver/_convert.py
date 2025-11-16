# Author: Timur Ulyahin, https://github.com/ucomru
# License: MIT – provided "as-is" without any warranty or liability
# Copyright: (c) 2024 Timur Ulyahin

"""
MODULE OVERVIEW:
    This module implements the low-level IPC-based `convert` function used to perform
    document format conversion by invoking platform-specific scripts for macOS (JXA)
    and Windows (PowerShell).

    It serves as the internal transport layer between Python and the operating-system
    automation scripts, enabling the conversion of Microsoft Word documents (.doc, .docx)
    and other supported formats into output formats such as PDF, RTF, TXT, HTML, and more.

    The function returns structured information (status, message, error_code) that
    is interpreted by the high-level `conver()` API, which converts these results
    into Python exceptions on failure.

USAGE:
    To use the `convert` function, specify the `input_path`, `output_path`, and optionally
    `keep_open` to indicate whether the application should remain open post-conversion.

RECOMMENDED PATHS:
    Use paths like "~/Downloads/" on macOS or "C:/Users/YourName/Downloads/" on Windows
    for both `input_path` and `output_path` to ensure permission consistency.

SETUP (macOS):
    1. Go to System Settings > Privacy & Security > Files and Folders.
    2. Under Microsoft Word or Microsoft Excel, enable access to "Downloads/" if required.

SUPPORTED FORMATS:
    The function supports conversion of the following file extensions:
    - Microsoft Word (.docx)    // Modern Word document format
    - Word 97-2003 (.doc)       // Legacy Word document format
    - PDF (.pdf)                // Export documents to PDF format
    - Rich Text Format (.rtf)   // Basic document format with limited formatting
    - OpenDocument Text (.odt)  // OpenOffice/LibreOffice document format
    - Plain Text (.txt)         // Plain text format without any formatting
    - HTML (.html)              // Web page format for viewing in browsers

STATUS SCHEMA (ConvertResult):
    The function returns a structured dictionary as follows:
    {
        "status": "success" | "error",        // Execution status
        "input": "<path to input file>",      // Path to the input file (may be null in error cases)
        "output": "<path to output file>",    // Path to the output file (may be null in error cases)
        "message": "OK" | "<error message>",  // "OK" on success or specific error message
        "error_code": 0 | <error code>        // 0 for success; specific code for different errors
    }

ERROR CODES:
    The `convert` function and associated scripts define a set of error codes for troubleshooting:
    0   - Success
    1   - Incorrect JSON format or missing required fields
    2   - Input file format is unsupported
    3   - Output file format is unsupported
    11  - Input file not found
    21  - Microsoft Word or Excel did not start within the expected time
    31  - Error saving or converting file
    98  - Invalid JSON output from script
    99  - Unsupported platform
"""

import subprocess
from json import loads, dumps, JSONDecodeError
import sys
from importlib.resources import files, as_file
from typing import TypedDict, Optional, List


class ConvertResult(TypedDict):
    status: str
    input: Optional[str]
    output: Optional[str]
    message: str
    error_code: int


def convert(
    input_path: str, output_path: str, keep_open: bool = False
) -> ConvertResult:
    """
    Convert a document from one format to another using platform-specific scripts.

    This function acts as an interface for inter-process communication (IPC) with macOS
    or Windows scripts to perform document conversions. Depending on the operating system,
    it triggers a JXA or PowerShell script, which performs the conversion and returns a
    structured JSON result. The `convert` function processes the JSON result, ensuring
    it matches the strict structure defined by `ConvertResult`.

    Parameters:
        input_path (str): The path to the source file.
        output_path (str): The path where the converted file will be saved.
        keep_open (bool): Whether to keep the application (Word/Excel) open after processing.

    Returns:
        ConvertResult: A dictionary with keys for `status`, `input`, `output`, `message`, and
        `error_code`, providing details of the conversion process and any errors encountered.

    Errors:
        Returns structured error responses as defined by the ConvertResult type. Unexpected errors
        that are not captured by the script will return error_code 98 for JSON parsing failures.
    """
    if sys.platform == "darwin":
        return _run_macos_script(input_path, output_path, keep_open)
    elif sys.platform == "win32":
        return _run_windows_script(input_path, output_path, keep_open)
    else:
        # Unsupported platform
        return {
            "status": "error",
            "input": input_path,
            "output": output_path,
            "message": "Unsupported platform.",
            "error_code": 99,
        }


def _run_macos_script(
    input_path: str, output_path: str, keep_open: bool
) -> ConvertResult:
    """Run macOS JXA script for conversion."""
    with as_file(files("conver.scripts").joinpath("convert.jxa")) as script_path:
        command = [
            "osascript",
            "-l",
            "JavaScript",
            str(script_path),
            dumps({"input": input_path, "output": output_path, "keepOpen": keep_open}),
        ]
        return _execute_command(command, input_path, output_path)


def _run_windows_script(
    input_path: str, output_path: str, keep_open: bool
) -> ConvertResult:
    """Run Windows PowerShell script for conversion."""
    with as_file(files("conver.scripts").joinpath("convert.ps1")) as script_path:
        command = [
            "powershell",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
            "-jsonArgs",
            dumps({"input": input_path, "output": output_path, "keepOpen": keep_open}),
        ]
        return _execute_command(command, input_path, output_path)


def _execute_command(
    command: List[str], input_path: str, output_path: str
) -> ConvertResult:
    """
    Executes the command for either macOS or Windows and parses the JSON output.

    Attempts to run the provided script command and expects JSON output for successful completion.
    In cases where the script exits with an error, the `error_code` from the script’s JSON response
    is returned directly to help identify specific issues. If JSON parsing fails, a general error
    code of 98 is returned.

    Parameters:
        command (List[str]): The command to be executed.
        input_path (str): The path to the input file, passed for error context.
        output_path (str): The path to the output file, passed for error context.

    Returns:
        ConvertResult: A structured dictionary with status, message, error code, input, and output.
    """
    result = subprocess.run(command, capture_output=True, text=True)

    # osascript places all output in stderr, including success
    raw = result.stdout.strip() or result.stderr.strip()

    try:
        data = loads(raw)
    except JSONDecodeError:
        return {
            "status": "error",
            "input": input_path,
            "output": output_path,
            "message": "Invalid JSON output from script.",
            "error_code": 98,
        }

    # canonical normalized result
    return {
        "status": data.get("status", "error"),
        "input": data.get("input", input_path),
        "output": data.get("output", output_path),
        "message": data.get("message", ""),
        "error_code": data.get("error_code", result.returncode),
    }
