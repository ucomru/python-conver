"""Tests for IPC dispatch and JSON response normalization, without OS scripts."""

# These tests intentionally exercise private transport helpers.
# pyright: reportPrivateUsage=false

import json
import subprocess
from types import SimpleNamespace
from typing import cast
from unittest.mock import Mock

import pytest

import conver._convert as ipc


def _success_result(input: str, output: str) -> ipc.ConvertResult:
    return {
        "status": "success",
        "input": input,
        "output": output,
        "message": "OK",
        "error_code": 0,
    }


def _run_result(
    stdout: str = "", stderr: str = "", code: int = 0
) -> subprocess.CompletedProcess[str]:
    return subprocess.CompletedProcess(
        args=["fake-script"], returncode=code, stdout=stdout, stderr=stderr
    )


def test_command_uses_capture_output_and_text_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    raw = json.dumps({"status": "success", "message": "OK", "error_code": 0})
    run = Mock(return_value=_run_result(stdout=raw))
    monkeypatch.setattr(subprocess, "run", run)
    command = ["fake-script", "--example"]

    actual = ipc._execute_command(command, "in.docx", "out.pdf")

    assert actual == {
        "status": "success",
        "input": "in.docx",
        "output": "out.pdf",
        "message": "OK",
        "error_code": 0,
    }
    run.assert_called_once_with(command, capture_output=True, text=True, check=False)


def test_stderr_is_used_when_stdout_is_empty(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    raw = json.dumps({"status": "success", "error_code": 0})
    monkeypatch.setattr(subprocess, "run", Mock(return_value=_run_result(stderr=raw)))

    assert ipc._execute_command(["fake"], "in", "out")["status"] == "success"


@pytest.mark.parametrize("raw", ["", "not-json", "[]", "null", '"str"', "42"])
def test_invalid_json_or_non_object_returns_ipc_error(
    monkeypatch: pytest.MonkeyPatch, raw: str
) -> None:
    monkeypatch.setattr(subprocess, "run", Mock(return_value=_run_result(stdout=raw)))

    result = ipc._execute_command(["fake"], "in.doc", "out.pdf")

    assert result == {
        "status": "error",
        "input": "in.doc",
        "output": "out.pdf",
        "message": "Invalid JSON output from script.",
        "error_code": 98,
    }


def test_nonzero_process_exit_cannot_be_overridden_by_json_success(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    raw = json.dumps({"status": "success", "error_code": 0})
    monkeypatch.setattr(
        subprocess, "run", Mock(return_value=_run_result(stdout=raw, code=1))
    )

    result = ipc._execute_command(["fake"], "in", "out")

    assert result["status"] == "error"
    assert result["error_code"] == 98
    assert "code 1" in result["message"]


def test_nonzero_exit_preserves_script_error_code(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    raw = json.dumps({"status": "error", "error_code": 31, "message": "save error"})
    monkeypatch.setattr(
        subprocess, "run", Mock(return_value=_run_result(stdout=raw, code=1))
    )

    result = ipc._execute_command(["fake"], "in", "out")

    assert result["error_code"] == 31
    assert result["message"] == "save error"


def test_missing_json_fields_returns_ipc_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        subprocess, "run", Mock(return_value=_run_result(stdout="{}", code=7))
    )

    result = ipc._execute_command(["fake"], "in", "out")

    assert result == {
        "status": "error",
        "input": "in",
        "output": "out",
        "message": "Invalid JSON output from script.",
        "error_code": 98,
    }


@pytest.mark.parametrize("platform", ["darwin", "win32"])
def test_dispatches_platform_to_corresponding_runner(
    monkeypatch: pytest.MonkeyPatch, platform: str
) -> None:
    monkeypatch.setattr(ipc, "sys", SimpleNamespace(platform=platform))
    mac = Mock(return_value={"status": "success", "error_code": 0})
    windows = Mock(return_value={"status": "success", "error_code": 0})
    monkeypatch.setattr(ipc, "_run_macos_script", mac)
    monkeypatch.setattr(ipc, "_run_windows_script", windows)

    result = ipc.convert("in.docx", "out.pdf", keep_open=True)

    assert result["status"] == "success"
    selected, skipped = (mac, windows) if platform == "darwin" else (windows, mac)
    selected.assert_called_once_with("in.docx", "out.pdf", True)
    skipped.assert_not_called()


def test_unsupported_platform_returns_99(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(ipc, "sys", SimpleNamespace(platform="linux"))

    result = ipc.convert("in.docx", "out.pdf")

    assert result["status"] == "error"
    assert result["error_code"] == 99
    assert result["input"] == "in.docx"
    assert result["output"] == "out.pdf"


def test_macos_command_and_payload(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[tuple[list[str], str, str]] = []

    def fake_execute(command: list[str], input: str, output: str) -> ipc.ConvertResult:
        calls.append((command, input, output))
        return _success_result(input, output)

    monkeypatch.setattr(ipc, "_execute_command", fake_execute)

    _ = ipc._run_macos_script("input.docx", "output.pdf", True)

    assert len(calls) == 1
    command, source, destination = calls[0]
    assert (source, destination) == ("input.docx", "output.pdf")
    assert command[:3] == ["osascript", "-l", "JavaScript"]
    assert command[3].endswith("convert.jxa")
    assert cast(dict[str, object], json.loads(command[4])) == {
        "input": "input.docx",
        "output": "output.pdf",
        "keepOpen": True,
    }


def test_windows_command_and_payload(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[tuple[list[str], str, str]] = []

    def fake_execute(command: list[str], input: str, output: str) -> ipc.ConvertResult:
        calls.append((command, input, output))
        return _success_result(input, output)

    monkeypatch.setattr(ipc, "_execute_command", fake_execute)

    _ = ipc._run_windows_script("input.docx", "output.pdf", False)

    assert len(calls) == 1
    command, source, destination = calls[0]
    assert (source, destination) == ("input.docx", "output.pdf")
    assert command[:4] == ["powershell", "-ExecutionPolicy", "Bypass", "-File"]
    assert command[4].endswith("convert.ps1")
    assert command[5] == "-jsonArgs"
    assert cast(dict[str, object], json.loads(command[6])) == {
        "input": "input.docx",
        "output": "output.pdf",
        "keepOpen": False,
    }
