"""Tests for the public Python API, without launching Microsoft Word."""

from importlib import import_module
from pathlib import Path
from unittest.mock import Mock

import pytest

from conver import (
    StagingError as PublicStagingError,
)
from conver import (
    conver as public_conver,
)
from conver import (
    conver_stage as public_conver_stage,
)
from conver.conver import (
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

# The package exports a function called `conver`, so import its module explicitly.
api = import_module("conver.conver")


def test_public_api_exports() -> None:
    assert public_conver is conver
    assert public_conver_stage is conver_stage
    assert PublicStagingError is StagingError


def test_filename_only_output_is_next_to_input(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    source = tmp_path / "docs" / "source.docx"
    source.parent.mkdir()
    _ = source.write_text("input")
    backend = Mock(return_value={"status": "success", "error_code": 0})
    monkeypatch.setattr(api, "convert", backend)

    result = conver(source, "result.pdf", keep_open=True)

    assert result == source.parent / "result.pdf"
    backend.assert_called_once_with(
        input=str(source), output=str(result), keep_open=True
    )


def test_relative_output_with_directory_is_relative_to_cwd(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.chdir(tmp_path)
    source = tmp_path / "docs" / "source.docx"
    source.parent.mkdir()
    source.touch()
    monkeypatch.setattr(
        api, "convert", Mock(return_value={"status": "success", "error_code": 0})
    )

    result = conver("docs/source.docx", "output/result.pdf")

    assert result == tmp_path / "output" / "result.pdf"


def test_tilde_paths_are_expanded(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setenv("USERPROFILE", str(tmp_path))
    source = tmp_path / "source.docx"
    source.touch()
    backend = Mock(return_value={"status": "success", "error_code": 0})
    monkeypatch.setattr(api, "convert", backend)

    result = conver("~/source.docx", "~/result.pdf")

    assert result == tmp_path / "result.pdf"
    backend.assert_called_once_with(
        input=str(source), output=str(tmp_path / "result.pdf"), keep_open=False
    )


def test_missing_input_raises_before_backend(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    backend = Mock()
    monkeypatch.setattr(api, "convert", backend)

    with pytest.raises(InputFileNotFound) as exc_info:
        _ = conver(tmp_path / "missing.docx", "result.pdf")

    assert exc_info.value.error_code == 11
    backend.assert_not_called()


@pytest.mark.parametrize(
    ("code", "expected"),
    [
        (1, IPCError),
        (2, UnsupportedFormat),
        (3, UnsupportedFormat),
        (11, InputFileNotFound),
        (21, WordStartError),
        (31, SaveError),
        (98, IPCError),
        (99, PlatformNotSupported),
        (777, IPCError),
    ],
)
def test_error_code_maps_to_exception(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    code: int,
    expected: type[ConverError],
) -> None:
    source = tmp_path / "source.docx"
    source.touch()
    monkeypatch.setattr(
        api,
        "convert",
        Mock(return_value={"status": "error", "error_code": code, "message": "fail"}),
    )

    with pytest.raises(expected) as exc_info:
        _ = conver(source, "result.pdf")

    assert exc_info.value.error_code == code
    assert "fail" in str(exc_info.value)


@pytest.mark.parametrize(
    "backend_result",
    [
        {"status": "error", "error_code": 0, "message": "failed"},
        {"status": "unexpected", "error_code": 0},
        {"error_code": 0},
        {"status": "error", "error_code": None},
    ],
)
def test_inconsistent_or_missing_success_fields_raise_ipc_error(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    backend_result: dict[str, object],
) -> None:
    source = tmp_path / "source.docx"
    source.touch()
    monkeypatch.setattr(api, "convert", Mock(return_value=backend_result))

    with pytest.raises(IPCError) as exc_info:
        _ = conver(source, "result.pdf")

    assert exc_info.value.error_code == 98
