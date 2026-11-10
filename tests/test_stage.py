"""Tests for the macOS staging flow, using real temporary filesystem operations."""

import shutil
from importlib import import_module
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from conver.conver import (
    InputFileNotFound,
    PlatformNotSupported,
    SaveError,
    StagingError,
    UnsupportedFormat,
    conver_stage,
)

# Keep a module reference only for patching the backend and platform.
api = import_module("conver.conver")


@pytest.fixture
def stage_env(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> tuple[Path, Path, Path]:
    # Do not change global sys.platform, which can interfere with pytest itself.
    monkeypatch.setattr(api, "sys", SimpleNamespace(platform="darwin"))
    source = tmp_path / "source.docx"
    _ = source.write_text("original")
    staging = tmp_path / "staging"
    staging.mkdir()
    destination = tmp_path / "result.pdf"
    return source, staging, destination


def _write_result(input: Path, output: Path, keep_open: bool = False) -> Path:
    _ = input, keep_open
    _ = output.write_text("converted")
    return output


def test_stage_converts_moves_and_cleans(
    monkeypatch: pytest.MonkeyPatch,
    stage_env: tuple[Path, Path, Path],
) -> None:
    source, staging, destination = stage_env
    calls: list[tuple[Path, Path, bool]] = []

    def fake_conver(input: Path, output: Path, keep_open: bool = False) -> Path:
        calls.append((input, output, keep_open))
        return _write_result(input, output, keep_open)

    monkeypatch.setattr(api, "conver", fake_conver)

    result = conver_stage(source, destination, keep_open=True, stage_dir=staging)

    assert result == destination
    assert destination.read_text() == "converted"
    assert source.read_text() == "original"
    assert list(staging.iterdir()) == []
    assert list(destination.parent.glob("._conver-temp-*")) == []
    assert len(calls) == 1
    actual_input, staged_output, actual_keep_open = calls[0]
    assert actual_input == source
    assert staged_output.parent == staging
    assert staged_output.name.startswith("_conver-temp-")
    assert staged_output.suffix == ".pdf"
    assert actual_keep_open is True


def test_stage_replaces_existing_destination(
    monkeypatch: pytest.MonkeyPatch,
    stage_env: tuple[Path, Path, Path],
) -> None:
    source, staging, destination = stage_env
    _ = destination.write_text("old result")
    monkeypatch.setattr(api, "conver", _write_result)

    assert conver_stage(source, destination, stage_dir=staging) == destination
    assert destination.read_text() == "converted"
    assert list(staging.iterdir()) == []


def test_stage_uses_downloads_by_default(
    monkeypatch: pytest.MonkeyPatch,
    stage_env: tuple[Path, Path, Path],
    tmp_path: Path,
) -> None:
    source, _, destination = stage_env
    downloads = tmp_path / "Downloads"
    downloads.mkdir()

    def fake_home(_cls: type[Path]) -> Path:
        return tmp_path

    monkeypatch.setattr(Path, "home", classmethod(fake_home))
    staged_outputs: list[Path] = []

    def fake_conver(input: Path, output: Path, keep_open: bool = False) -> Path:
        staged_outputs.append(output)
        return _write_result(input, output, keep_open)

    monkeypatch.setattr(api, "conver", fake_conver)

    _ = conver_stage(source, destination)

    assert len(staged_outputs) == 1
    assert staged_outputs[0].parent == downloads
    assert list(downloads.iterdir()) == []


def test_stage_expands_tilde_directory(
    monkeypatch: pytest.MonkeyPatch,
    stage_env: tuple[Path, Path, Path],
    tmp_path: Path,
) -> None:
    source, _, destination = stage_env

    def fake_home(_cls: type[Path]) -> Path:
        return tmp_path

    monkeypatch.setattr(Path, "home", classmethod(fake_home))
    monkeypatch.setattr(api, "conver", _write_result)

    _ = conver_stage(source, destination, stage_dir="~")

    assert destination.read_text() == "converted"
    assert not list(tmp_path.glob("_conver-temp-*"))


@pytest.mark.parametrize("platform", ["win32", "linux"])
def test_stage_is_macos_only(
    monkeypatch: pytest.MonkeyPatch,
    stage_env: tuple[Path, Path, Path],
    platform: str,
) -> None:
    source, staging, destination = stage_env
    monkeypatch.setattr(api, "sys", SimpleNamespace(platform=platform))

    with pytest.raises(PlatformNotSupported) as exc_info:
        _ = conver_stage(source, destination, stage_dir=staging)

    assert exc_info.value.error_code == 99


def test_missing_source_is_rejected(
    stage_env: tuple[Path, Path, Path],
) -> None:
    source, staging, destination = stage_env
    source.unlink()
    with pytest.raises(InputFileNotFound):
        _ = conver_stage(source, destination, stage_dir=staging)


@pytest.mark.parametrize("extension", [".html", ".htm", ".HTML", ".HTM"])
def test_html_output_is_rejected(
    stage_env: tuple[Path, Path, Path], extension: str
) -> None:
    source, staging, _ = stage_env
    with pytest.raises(UnsupportedFormat) as exc_info:
        _ = conver_stage(source, source.with_suffix(extension), stage_dir=staging)
    assert exc_info.value.error_code == 3


def test_missing_staging_directory_is_rejected(
    stage_env: tuple[Path, Path, Path],
) -> None:
    source, staging, destination = stage_env
    with pytest.raises(StagingError, match="Staging directory does not exist"):
        _ = conver_stage(source, destination, stage_dir=staging / "missing")


def test_staging_cannot_overwrite_input(
    monkeypatch: pytest.MonkeyPatch,
    stage_env: tuple[Path, Path, Path],
) -> None:
    source, staging, _ = stage_env
    backend = Mock()
    monkeypatch.setattr(api, "conver", backend)

    with pytest.raises(StagingError, match="cannot overwrite the input"):
        _ = conver_stage(source, source, stage_dir=staging)

    backend.assert_not_called()


def test_conversion_error_preserves_exception_and_cleans_staged_file(
    monkeypatch: pytest.MonkeyPatch,
    stage_env: tuple[Path, Path, Path],
) -> None:
    source, staging, destination = stage_env

    def fail_conversion(input: Path, output: Path, keep_open: bool = False) -> Path:
        _ = input, keep_open
        _ = output.write_text("partial result")
        raise SaveError("Cannot save", 31)

    monkeypatch.setattr(api, "conver", fail_conversion)

    with pytest.raises(SaveError) as exc_info:
        _ = conver_stage(source, destination, stage_dir=staging)

    assert exc_info.value.error_code == 31
    assert list(staging.iterdir()) == []
    assert not destination.exists()


def test_move_error_becomes_staging_error_and_cleans(
    monkeypatch: pytest.MonkeyPatch,
    stage_env: tuple[Path, Path, Path],
) -> None:
    source, staging, destination = stage_env
    monkeypatch.setattr(api, "conver", _write_result)
    monkeypatch.setattr(shutil, "move", Mock(side_effect=OSError("permission denied")))

    with pytest.raises(StagingError, match="permission denied") as exc_info:
        _ = conver_stage(source, destination, stage_dir=staging)

    assert isinstance(exc_info.value.__cause__, OSError)
    assert list(staging.iterdir()) == []
    assert not destination.exists()


def test_final_replace_error_cleans_pending_file(
    monkeypatch: pytest.MonkeyPatch,
    stage_env: tuple[Path, Path, Path],
) -> None:
    source, staging, destination = stage_env
    destination.mkdir()  # Replacing a directory with a file must fail.
    monkeypatch.setattr(api, "conver", _write_result)

    with pytest.raises(StagingError):
        _ = conver_stage(source, destination, stage_dir=staging)

    assert destination.is_dir()
    assert list(staging.iterdir()) == []
    assert not list(destination.parent.glob("._conver-temp-*"))
