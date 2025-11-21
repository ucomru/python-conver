import sys
from pathlib import Path
from typing import Union

from click import (
    command,
    argument,
    option,
    version_option,
    echo,
    Path as ClickPath,
    UsageError,
)

from .conver import conver, ConverError
from .__version__ import __version__


def fail(label: str, err: Exception, code: int = 1):
    echo(f"{label} {err}", err=True)
    sys.exit(code)


def _infer_common_parent(paths: list[Path]) -> Union[Path, None]:
    parents = {p.resolve().parent for p in paths}
    return parents.pop() if len(parents) == 1 else None


@command(help="Convert Word documents.")
@argument(
    "inputs", type=ClickPath(exists=True, path_type=Path), nargs=-1, metavar="INPUT..."
)
@option(
    "-o",
    "--output",
    "output",
    type=ClickPath(path_type=Path),
    required=False,
    help="Output file or directory (for multiple inputs).",
)
@option("-p", "--pdf", "target", flag_value="pdf", help="Convert to PDF.")
@option("-d", "--docx", "target", flag_value="docx", help="Convert to DOCX.")
@option("--doc", "target", flag_value="doc", help="Convert to DOC.")
@option("-r", "--rtf", "target", flag_value="rtf", help="Convert to RTF.")
@option("--odt", "target", flag_value="odt", help="Convert to ODT.")
@option("-t", "--txt", "target", flag_value="txt", help="Convert to TXT.")
@option("-h", "--html", "target", flag_value="html", help="Convert to HTML.")
@option("-k", "--keep-open", is_flag=True, help="Keep Microsoft Word open.")
@version_option(__version__, "-V", "--version")
def cli(inputs, output, target, keep_open):
    """Main CLI entry point."""
    if not inputs:
        raise UsageError("No input files specified.")

    if output is not None and target is not None:
        raise UsageError("Cannot use both --output and format flags together.")

    # --- MULTIPLE INPUTS ---
    if len(inputs) > 1:
        if output is None:
            output = _infer_common_parent(inputs)

        if output is None:
            raise UsageError(
                "Input files are in different directories; specify --output DIRECTORY."
            )

        output = output.resolve()

        if output.exists() and not output.is_dir():
            raise UsageError("--output must be a directory for multiple inputs.")

        output.mkdir(parents=True, exist_ok=True)

        if target is None:
            target = "pdf"

        for inp in inputs:
            out_file = output / (inp.stem + f".{target}")
            try:
                result = conver(inp, out_file, keep_open=keep_open)
                echo(result)
            except ConverError as err:
                fail("Error:", err, err.error_code or 1)

    # --- SINGLE INPUT ---
    else:
        inp = inputs[0]

        if output is not None:
            try:
                result = conver(inp, output, keep_open=keep_open)
                echo(result)
            except ConverError as err:
                fail("Error:", err, err.error_code or 1)
        else:
            if target is None:
                target = "pdf"

            out_file = inp.with_suffix("." + target)
            try:
                result = conver(inp, out_file, keep_open=keep_open)
                echo(result)
            except ConverError as err:
                fail("Error:", err, err.error_code or 1)
