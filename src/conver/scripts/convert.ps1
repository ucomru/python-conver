<#
Author: Timur Ulyahin, https://github.com/ucomru
License: MIT – provided "as-is" without any warranty or liability
Copyright: (c) 2024 Timur Ulyahin

USAGE:
    powershell -ExecutionPolicy Bypass -File convert.ps1 -jsonArgs `
        "{\"input\": \"<inputPath>\", \"output\": \"<outputPath>\", \"keepOpen\": true | false}"

RECOMMENDED PATHS:
    Use paths like "C:\Users\UserName\Downloads\" for both <inputPath> and <outputPath>.

SETUP (Windows):
    1. Ensure Microsoft Word is installed on your system.
    2. If prompted, allow PowerShell to automate Microsoft Word.

SUPPORTED FORMATS:
    The script supports the following file extensions for input and output:
    - PDF (.pdf)                # Export documents to PDF format
    - Microsoft Word (.docx)    # Modern Word document format
    - Word 97-2003 (.doc)       # Legacy Word document format
    - Rich Text Format (.rtf)   # Basic document format with limited formatting
    - Plain Text (.txt)         # Plain text format without any formatting
    - HTML (.html, .htm)        # Web page format for viewing in browsers
    - OpenDocument Text (.odt)  # OpenOffice/LibreOffice document format

PARAMETERS:
    - input: "<path to input file>"     # Path to the input file (required)
    - output: "<path to output file>"   # Path to the output file (required)
    - keepOpen: true | false            # Optional. Keep Word open (default: false)

STATUS SCHEMA (JSON Output):
    The script outputs a JSON object with the following fields:
    {
        "status": "success" | "error",        # Execution status
        "input": "<path to input file>",      # Path to the input file
        "output": "<path to output file>",    # Path to the output file (null if error)
        "message": "OK" | "<error message>",  # "OK" on success or specific error message
        "error_code": 0 | <error code>        # 0 for success; specific code for different errors
    }

ERROR CODES:
    0   - Success
    1   - Incorrect JSON format or missing required fields
    2   - Input file format is unsupported
    3   - Output file format is unsupported
    11  - Input file not found
    21  - Microsoft Word did not start or is not installed
    31  - Error saving or converting file
#>

param (
    [string]$jsonArgs
)

# Supported file formats and corresponding codes for Microsoft Word
$formatCodes = @{
    "pdf" = 17
    "docx" = 16
    "doc" = 0
    "rtf" = 6
    "txt" = 7
    "html" = 8
    "htm" = 8
    "odt" = 19
}

# Function to retrieve format code by extension
function Get-FormatCodeByExtension {
    param ($extension)
    $extension = $extension.ToLower()
    return $formatCodes[$extension]
}

# Parse JSON input
try {
    $params = $jsonArgs | ConvertFrom-Json
} catch {
    Write-Output (ConvertTo-Json @{
        status = "error"
        input = $null
        output = $null
        message = "Invalid JSON format."
        error_code = 1
    })
    exit 1
}

$inputPath = $params.input
$outputPath = $params.output
$keepOpen = $params.keepOpen -eq $true

# Validate required parameters
if (-not $inputPath -or -not $outputPath) {
    Write-Output (ConvertTo-Json @{
        status = "error"
        input = $inputPath
        output = $outputPath
        message = "Both 'input' and 'output' fields are required in JSON."
        error_code = 1
    })
    exit 1
}

# Check input file format support
$inputExtension = [System.IO.Path]::GetExtension($inputPath).TrimStart('.').ToLower()
$inputFormat = Get-FormatCodeByExtension -extension $inputExtension

if ($null -eq $inputFormat) {
    Write-Output (ConvertTo-Json @{
        status = "error"
        input = $inputPath
        output = $null
        message = "The input file format '$inputExtension' is unsupported."
        error_code = 2
    })
    exit 2
}

# Check output file format support
$outputExtension = [System.IO.Path]::GetExtension($outputPath).TrimStart('.').ToLower()
$outputFormat = Get-FormatCodeByExtension -extension $outputExtension

if ($null -eq $outputFormat) {
    Write-Output (ConvertTo-Json @{
        status = "error"
        input = $inputPath
        output = $outputPath
        message = "The output file format '$outputExtension' is unsupported."
        error_code = 3
    })
    exit 3
}

# Initialize Word COM object with error handling
try {
    $word = New-Object -ComObject Word.Application -ErrorAction Stop
} catch {
    Write-Output (ConvertTo-Json @{
        status = "error"
        input = $inputPath
        output = $outputPath
        message = "Microsoft Word is not installed or cannot be started."
        error_code = 21
    })
    exit 21
}

$doc = $null

try {
    $word.Visible = $false

    # Check if input file exists
    if (-not (Test-Path -Path $inputPath)) {
        Write-Output (ConvertTo-Json @{
            status = "error"
            input = $inputPath
            output = $null
            message = "File '$inputPath' not found."
            error_code = 11
        })
        exit 11
    }

    # Open and convert the document
    $doc = $word.Documents.Open($inputPath)
    $doc.SaveAs([ref]$outputPath, [ref]$outputFormat)
    $doc.Close([ref]0)
    $doc = $null
} catch {
    Write-Output (ConvertTo-Json @{
        status = "error"
        input = $inputPath
        output = $outputPath
        message = $_.Exception.Message
        error_code = 31
    })
    exit 31
} finally {
    # Close Word if opened by the script and keepOpen is false
    if ($null -ne $doc) {
        try { $doc.Close([ref]0) } catch {}
    }

    if (-not $keepOpen) {
        try { $word.Quit() } catch {}
    }
}

# Success message
Write-Output (ConvertTo-Json @{
    status = "success"
    input = $inputPath
    output = $outputPath
    message = "OK"
    error_code = 0
})
