import sys
from collections.abc import Iterable
from pathlib import Path

from click import (
    Path as ClickPath,
)
from click import (
    UsageError,
    argument,
    command,
    echo,
    option,
    version_option,
)

from .__version__ import __version__
from .conver import ConverError, conver, conver_stage

WORD_EXTENSIONS = frozenset({".docx", ".doc", ".rtf", ".txt", ".html", ".htm", ".odt"})


def _fail(label: str, err: Exception, code: int = 1):
    echo(f"{label} {err}", err=True)
    sys.exit(code)


def _is_not_word_temp_file(path: Path) -> bool:
    return not (path.name.startswith("~$") and path.suffix.lower() in WORD_EXTENSIONS)


def _is_supported_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in WORD_EXTENSIONS


def _infer_common_parent(paths: Iterable[Path]) -> Path | None:
    parents = {p.resolve().parent for p in paths}
    return parents.pop() if len(parents) == 1 else None


@command(context_settings={"help_option_names": ["-h", "--help"]})
@argument(
    "inputs", type=ClickPath(exists=True, path_type=Path), nargs=-1, metavar="INPUT..."
)
@option(
    "-o",
    "--output",
    "output",
    type=ClickPath(path_type=Path),
    help="Output file or directory.",
)
@option("-p", "--pdf", "target", flag_value="pdf", help="Convert to PDF [DEFAULT]")
@option("-x", "--docx", "target", flag_value="docx", help="Convert to DOCX")
@option("-d", "--doc", "target", flag_value="doc", help="Convert to DOC")
@option("-r", "--rtf", "target", flag_value="rtf", help="Convert to RTF")
@option("-t", "--txt", "target", flag_value="txt", help="Convert to TXT")
@option("-m", "--html", "target", flag_value="html", help="Convert to HTML")
@option("-u", "--odt", "target", flag_value="odt", help="Convert to ODT")
@option("-k", "--keep-open", is_flag=True, help="Keep Microsoft Word open")
@option("-q", "--quiet", is_flag=True, help="Suppress output on success")
@option("-e", "--ext-only", is_flag=True, help="Supported extensions only")
@option(
    "-s",
    "--stage/--no-stage",
    default=False,
    envvar="CONVER_STAGE",
    help="Use temporary output staging (macOS only)",
)
@option(
    "--stage-dir",
    type=ClickPath(path_type=Path),
    default=Path.home() / "Downloads",
    envvar="CONVER_STAGE_DIR",
    help="Directory for staged output [~/Downloads]",
)
@option("--include-temp", is_flag=True, help="Include MS Word temp files ~$...")
@version_option(__version__, "-v", "--version")
def cli(
    inputs: tuple[Path, ...],
    output: Path | None,
    target: str | None,
    keep_open: bool,
    quiet: bool,
    ext_only: bool,
    stage: bool,
    stage_dir: Path,
    include_temp: bool,
) -> None:
    """Convert documents using Microsoft Word."""
    if not include_temp:
        inputs = tuple(filter(_is_not_word_temp_file, inputs))

    if ext_only:
        inputs = tuple(filter(_is_supported_file, inputs))

    if not inputs:
        raise UsageError("No input files specified.")

    if stage:
        if sys.platform != "darwin":
            raise UsageError("Staging mode is supported only on macOS.")

        stage_dir = stage_dir.expanduser().resolve()

        if not stage_dir.is_dir():
            raise UsageError(f"Staging directory does not exist: {stage_dir}")

    # --- MULTIPLE INPUTS ---
    if len(inputs) > 1:
        if output is None:
            output = _infer_common_parent(inputs)

        if output is None:
            raise UsageError("In batch mode, --output DIRECTORY is required.")

        output = output.resolve()

        if output.exists() and not output.is_dir():
            raise UsageError("--output must be a directory for multiple inputs.")

        output.mkdir(parents=True, exist_ok=True)

        if target is None:
            target = "pdf"

        for index, input in enumerate(inputs):
            # Keep Word open for all but the last file
            effective_keep_open = (index < len(inputs) - 1) or keep_open

            output_file = output / (input.stem + f".{target}")

            try:
                if not stage or output_file.suffix.lower() in {".html", ".htm"}:
                    result = conver(
                        input=input,
                        output=output_file,
                        keep_open=effective_keep_open,
                    )
                else:
                    result = conver_stage(
                        input=input,
                        output=output_file,
                        keep_open=effective_keep_open,
                        stage_dir=stage_dir,
                    )

                if not quiet:
                    echo(result)
            except ConverError as err:
                _fail("Error:", err, err.error_code or 1)

    # --- SINGLE INPUT ---
    elif len(inputs) == 1:
        input = inputs[0]

        if output is not None:
            output = output.resolve()

            if output.suffix:
                if target is not None:
                    raise UsageError(
                        "Cannot use format flag when output is a file path."
                    )
                output_file = output
            else:
                if target is None:
                    target = "pdf"

                output.mkdir(parents=True, exist_ok=True)
                output_file = output / (input.stem + f".{target}")

        else:
            if target is None:
                target = "pdf"

            output_file = input.with_suffix("." + target)

        try:
            if not stage or output_file.suffix.lower() in {".html", ".htm"}:
                result = conver(
                    input=input,
                    output=output_file,
                    keep_open=keep_open,
                )
            else:
                result = conver_stage(
                    input=input,
                    output=output_file,
                    keep_open=keep_open,
                    stage_dir=stage_dir,
                )

            if not quiet:
                echo(result)
        except ConverError as err:
            _fail("Error:", err, err.error_code or 1)
