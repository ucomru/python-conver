"""Click CLI tests: argument parsing, filtering, batching, and staging dispatch."""

import sys
from importlib import import_module
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from click.testing import CliRunner

from conver.cli import cli
from conver.conver import SaveError

cli_module = import_module("conver.cli")


@pytest.fixture(autouse=True)
def clean_stage_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("CONVER_STAGE", raising=False)
    monkeypatch.delenv("CONVER_STAGE_DIR", raising=False)


@pytest.fixture
def calls(monkeypatch: pytest.MonkeyPatch) -> list[tuple[Path, Path, bool]]:
    records: list[tuple[Path, Path, bool]] = []

    def fake_conver(*, input: Path, output: Path, keep_open: bool = False) -> Path:
        records.append((input, output, keep_open))
        return output

    monkeypatch.setattr(cli_module, "conver", fake_conver)
    return records


def _touch(path: Path) -> Path:
    _ = path.write_text("input")
    return path


def test_default_pdf_and_output_message(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    source = _touch(tmp_path / "report.docx")

    result = CliRunner().invoke(cli, [str(source)])

    assert result.exit_code == 0, result.output
    assert calls == [(source, source.with_suffix(".pdf"), False)]
    assert str(source.with_suffix(".pdf")) in result.output


@pytest.mark.parametrize(
    ("flag", "extension"),
    [
        ("-p", ".pdf"),
        ("-x", ".docx"),
        ("-d", ".doc"),
        ("-r", ".rtf"),
        ("-t", ".txt"),
        ("-m", ".html"),
        ("-u", ".odt"),
    ],
)
def test_format_flags(
    tmp_path: Path,
    calls: list[tuple[Path, Path, bool]],
    flag: str,
    extension: str,
) -> None:
    source = _touch(tmp_path / "report.rtf")

    result = CliRunner().invoke(cli, [str(source), flag, "-q"])

    assert result.exit_code == 0, result.output
    assert calls[0][1] == source.with_suffix(extension)
    assert result.output == ""


def test_explicit_output_file(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    source = _touch(tmp_path / "input.docx")
    destination = tmp_path / "result.rtf"

    result = CliRunner().invoke(cli, [str(source), "-o", str(destination)])

    assert result.exit_code == 0, result.output
    assert calls == [(source, destination, False)]


def test_single_input_output_directory_is_created(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    source = _touch(tmp_path / "input.docx")
    outdir = tmp_path / "converted"

    result = CliRunner().invoke(cli, [str(source), "-o", str(outdir), "-q"])

    assert result.exit_code == 0, result.output
    assert outdir.is_dir()
    assert calls == [(source, outdir / "input.pdf", False)]


def test_format_flag_cannot_be_combined_with_output_filename(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    source = _touch(tmp_path / "input.docx")

    result = CliRunner().invoke(
        cli, [str(source), "-o", str(tmp_path / "result.pdf"), "-r"]
    )

    assert result.exit_code == 2
    assert "Cannot use format flag" in result.output
    assert calls == []


def test_ext_only_filters_unsupported_files_and_directories(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    docx = _touch(tmp_path / "one.DOCX")
    rtf = _touch(tmp_path / "two.rtf")
    png = _touch(tmp_path / "image.png")
    folder = tmp_path / "folder.docx"
    folder.mkdir()

    result = CliRunner().invoke(
        cli, [str(docx), str(png), str(folder), str(rtf), "-e", "-q"]
    )

    assert result.exit_code == 0, result.output
    assert [record[0] for record in calls] == [docx, rtf]


def test_temp_word_files_are_excluded_by_default(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    temp = _touch(tmp_path / "~$draft.docx")
    regular = _touch(tmp_path / "regular.docx")

    result = CliRunner().invoke(cli, [str(temp), str(regular), "-e", "-q"])

    assert result.exit_code == 0, result.output
    assert [record[0] for record in calls] == [regular]


def test_include_temp_overrides_temp_filter(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    temp = _touch(tmp_path / "~$draft.docx")

    result = CliRunner().invoke(cli, [str(temp), "-e", "--include-temp", "-q"])

    assert result.exit_code == 0, result.output
    assert [record[0] for record in calls] == [temp]


def test_ext_only_with_no_matching_files_is_usage_error(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    unsupported = _touch(tmp_path / "image.png")

    result = CliRunner().invoke(cli, [str(unsupported), "-e"])

    assert result.exit_code == 2
    assert calls == []


def test_batch_infers_common_parent_and_keep_open(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    sources = [_touch(tmp_path / f"input_{index}.docx") for index in range(3)]

    result = CliRunner().invoke(cli, [*(str(path) for path in sources), "-q"])

    assert result.exit_code == 0, result.output
    assert calls == [
        (sources[0], tmp_path / "input_0.pdf", True),
        (sources[1], tmp_path / "input_1.pdf", True),
        (sources[2], tmp_path / "input_2.pdf", False),
    ]


def test_keep_open_flag_applies_to_last_batch_file(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    a = _touch(tmp_path / "a.docx")
    b = _touch(tmp_path / "b.docx")

    result = CliRunner().invoke(cli, [str(a), str(b), "-k", "-q"])

    assert result.exit_code == 0, result.output
    assert [entry[2] for entry in calls] == [True, True]


def test_batch_from_different_directories_requires_output(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    left = tmp_path / "left"
    right = tmp_path / "right"
    left.mkdir()
    right.mkdir()
    a = _touch(left / "a.docx")
    b = _touch(right / "b.docx")

    result = CliRunner().invoke(cli, [str(a), str(b)])

    assert result.exit_code == 2
    assert "--output DIRECTORY is required" in result.output
    assert calls == []


def test_batch_creates_output_directory(
    tmp_path: Path, calls: list[tuple[Path, Path, bool]]
) -> None:
    a = _touch(tmp_path / "a.docx")
    b = _touch(tmp_path / "b.docx")
    outdir = tmp_path / "converted"

    result = CliRunner().invoke(cli, [str(a), str(b), "-o", str(outdir), "-q"])

    assert result.exit_code == 0, result.output
    assert outdir.is_dir()
    assert [entry[1].parent for entry in calls] == [outdir, outdir]


def test_backend_error_exit_code_and_message(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    source = _touch(tmp_path / "input.docx")
    monkeypatch.setattr(
        cli_module, "conver", Mock(side_effect=SaveError("Cannot save", 31))
    )

    result = CliRunner().invoke(cli, [str(source)])

    assert result.exit_code == 31
    assert "Cannot save" in result.output


def test_staging_requires_macos(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    source = _touch(tmp_path / "input.docx")
    monkeypatch.setattr(
        cli_module, "sys", SimpleNamespace(platform="win32", exit=sys.exit)
    )

    result = CliRunner().invoke(cli, [str(source), "-s"])

    assert result.exit_code == 2
    assert "only on macOS" in result.output


def test_stage_dispatches_to_stage_api(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    calls: list[tuple[Path, Path, bool]],
) -> None:
    source = _touch(tmp_path / "input.docx")
    staging = tmp_path / "staging"
    staging.mkdir()
    monkeypatch.setattr(
        cli_module, "sys", SimpleNamespace(platform="darwin", exit=sys.exit)
    )
    staged = Mock(return_value=source.with_suffix(".pdf"))
    monkeypatch.setattr(cli_module, "conver_stage", staged)

    result = CliRunner().invoke(
        cli,
        [str(source), "--stage", "--stage-dir", str(staging), "-q"],
    )

    assert result.exit_code == 0, result.output
    staged.assert_called_once_with(
        input=source,
        output=source.with_suffix(".pdf"),
        keep_open=False,
        stage_dir=staging,
    )
    assert calls == []


def test_html_output_bypasses_stage(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    calls: list[tuple[Path, Path, bool]],
) -> None:
    source = _touch(tmp_path / "input.docx")
    staging = tmp_path / "staging"
    staging.mkdir()
    monkeypatch.setattr(
        cli_module, "sys", SimpleNamespace(platform="darwin", exit=sys.exit)
    )
    staged = Mock()
    monkeypatch.setattr(cli_module, "conver_stage", staged)

    result = CliRunner().invoke(
        cli,
        [str(source), "-m", "-s", "--stage-dir", str(staging), "-q"],
    )

    assert result.exit_code == 0, result.output
    assert calls == [(source, source.with_suffix(".html"), False)]
    staged.assert_not_called()


def test_no_stage_overrides_environment_variable(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    calls: list[tuple[Path, Path, bool]],
) -> None:
    source = _touch(tmp_path / "input.docx")
    staging = tmp_path / "staging"
    staging.mkdir()
    monkeypatch.setenv("CONVER_STAGE", "1")
    monkeypatch.setattr(
        cli_module, "sys", SimpleNamespace(platform="darwin", exit=sys.exit)
    )
    staged = Mock()
    monkeypatch.setattr(cli_module, "conver_stage", staged)

    result = CliRunner().invoke(
        cli,
        [str(source), "--no-stage", "--stage-dir", str(staging), "-q"],
    )

    assert result.exit_code == 0, result.output
    assert calls == [(source, source.with_suffix(".pdf"), False)]
    staged.assert_not_called()
