from __future__ import annotations

import subprocess
from types import SimpleNamespace
from unittest.mock import patch

import hwp_com


class FakePywin32ImportError(ImportError):
    pass


class FakeComStartupFailure(OSError):
    pass


class FakeWindow:
    def __init__(self) -> None:
        self.Visible = False


class FakeWindows:
    def __init__(self) -> None:
        self.window = FakeWindow()

    def Item(self, index: int) -> FakeWindow:
        return self.window


class FakeHwpWithMessageMode:
    def __init__(self) -> None:
        self.XHwpWindows = FakeWindows()
        self.message_modes: list[int] = []
        self.registered: list[tuple[str, str]] = []
        self.quit_called = False

    def SetMessageBoxMode(self, mode: int) -> None:
        self.message_modes.append(mode)

    def RegisterModule(self, module_name: str, module_type: str) -> None:
        self.registered.append((module_name, module_type))

    def Quit(self) -> None:
        self.quit_called = True


class FakeHwpWithoutMessageMode:
    def __init__(self) -> None:
        self.XHwpWindows = FakeWindows()
        self.registered: list[tuple[str, str]] = []
        self.quit_called = False

    def RegisterModule(self, module_name: str, module_type: str) -> None:
        self.registered.append((module_name, module_type))

    def Quit(self) -> None:
        self.quit_called = True


def test_create_hwp_object_sets_message_box_mode_when_available() -> None:
    fake_hwp = FakeHwpWithMessageMode()
    fake_client = SimpleNamespace(Dispatch=lambda name: fake_hwp)

    def fake_import_module(name: str):
        if name == "win32com.client":
            return fake_client
        raise ImportError(name)

    warnings: list[str] = []
    with patch.object(hwp_com.importlib, "import_module", side_effect=fake_import_module):
        hwp = hwp_com.create_hwp_object(visible=True, startup_warnings=warnings)

    assert hwp is fake_hwp
    assert fake_hwp.message_modes == [0]
    assert fake_hwp.registered == [("FilePathCheckDLL", "SecurityModule")]
    assert fake_hwp.XHwpWindows.window.Visible is True
    assert warnings == []


def test_preflight_worker_warns_when_message_box_mode_is_unavailable(capsys) -> None:
    fake_hwp = FakeHwpWithoutMessageMode()
    fake_client = SimpleNamespace(Dispatch=lambda name: fake_hwp)

    def fake_import_module(name: str):
        if name == "win32com.client":
            return fake_client
        raise ImportError(name)

    with patch.object(hwp_com.importlib, "import_module", side_effect=fake_import_module):
        result = hwp_com.run_hwp_preflight_worker(visible=False)

    captured = capsys.readouterr()
    assert result == "HWP COM preflight OK: HWPFrame.HwpObject 생성 및 SecurityModule 등록 성공"
    assert "SetMessageBoxMode를 사용할 수 없어" in captured.err
    assert fake_hwp.quit_called


def test_startup_error_message_lists_actionable_checks() -> None:
    message = hwp_com.format_hwp_startup_error(FakeComStartupFailure("COM failure"))

    assert "pywin32 설치" in message
    assert "Hancom Office HWP 설치" in message
    assert "HWPFrame.HwpObject COM 등록" in message
    assert "FilePathCheckDLL SecurityModule 등록" in message
    assert "gen_py 캐시" in message


def test_create_hwp_object_raises_missing_pywin32_error_when_import_fails() -> None:
    def fake_import_module(name: str):
        if name == "win32com.client":
            raise FakePywin32ImportError(name)
        raise ImportError(name)

    error: hwp_com.MissingPywin32Error | None = None
    with patch.object(hwp_com.importlib, "import_module", side_effect=fake_import_module):
        try:
            hwp_com.create_hwp_object()
        except hwp_com.MissingPywin32Error as exc:
            error = exc

    assert error is not None
    message = str(error)
    assert "pywin32" in message
    assert "pip install pywin32" in message


def test_create_hwp_object_raises_startup_error_when_dispatch_fails() -> None:
    def dispatch_failure(name: str) -> None:
        raise FakeComStartupFailure(f"{name} COM failure")

    fake_client = SimpleNamespace(Dispatch=dispatch_failure)

    def fake_import_module(name: str):
        if name == "win32com.client":
            return fake_client
        raise ImportError(name)

    error: hwp_com.HwpStartupError | None = None
    with patch.object(hwp_com.importlib, "import_module", side_effect=fake_import_module):
        try:
            hwp_com.create_hwp_object()
        except hwp_com.HwpStartupError as exc:
            error = exc

    assert error is not None
    message = str(error)
    assert "HWP COM 자동화 시작 실패" in message
    assert "HWPFrame.HwpObject COM failure" in message


def test_preflight_timeout_includes_partial_worker_output() -> None:
    timeout = subprocess.TimeoutExpired(
        cmd=["python", "to_hwpx_com.py", "--_preflight-worker"],
        timeout=45,
        output="worker stdout",
        stderr="[HWP-PREFLIGHT] dispatch HWPFrame.HwpObject",
    )

    error: hwp_com.HwpPreflightTimeoutError | None = None
    with (
        patch.object(hwp_com, "running_hwp_pids", side_effect=[{100}, {100, 200}]),
        patch.object(hwp_com, "terminate_hwp_pids", return_value=["taskkill PID 200 exit=0: terminated"]),
        patch.object(hwp_com.subprocess, "run", side_effect=timeout),
    ):
        try:
            hwp_com.run_hwp_preflight(timeout=45)
        except hwp_com.HwpPreflightTimeoutError as exc:
            error = exc

    assert error is not None
    message = str(error)
    assert "timed out after 45 seconds" in message
    assert "worker stdout" in message
    assert "dispatch HWPFrame.HwpObject" in message
    assert "taskkill PID 200 exit=0" in message


def assert_preflight_worker_error_uses_output(
    stdout: str,
    stderr: str,
    expected: str,
) -> None:
    result = subprocess.CompletedProcess(
        args=["python", "to_hwpx_com.py", "--_preflight-worker"],
        returncode=1,
        stdout=stdout,
        stderr=stderr,
    )

    error: hwp_com.HwpPreflightWorkerError | None = None
    with (
        patch.object(hwp_com, "running_hwp_pids", return_value=set()),
        patch.object(hwp_com.subprocess, "run", return_value=result),
    ):
        try:
            hwp_com.run_hwp_preflight()
        except hwp_com.HwpPreflightWorkerError as exc:
            error = exc

    assert error is not None
    assert expected in str(error)


def test_preflight_nonzero_worker_raises_worker_error_with_stdout_fallback() -> None:
    assert_preflight_worker_error_uses_output(
        stdout="worker stdout fallback",
        stderr="",
        expected="worker stdout fallback",
    )


def test_preflight_nonzero_worker_raises_worker_error_with_stderr() -> None:
    assert_preflight_worker_error_uses_output(
        stdout="worker stdout",
        stderr="worker stderr failure",
        expected="worker stderr failure",
    )


def test_preflight_timeout_does_not_kill_preexisting_hwp_pids() -> None:
    timeout = subprocess.TimeoutExpired(
        cmd=["python", "to_hwpx_com.py", "--_preflight-worker"],
        timeout=45,
        stderr="[HWP-PREFLIGHT] dispatch HWPFrame.HwpObject",
    )

    error: hwp_com.HwpPreflightTimeoutError | None = None
    with (
        patch.object(hwp_com, "running_hwp_pids", side_effect=[{100}, {100, 200, 300}]),
        patch.object(hwp_com, "terminate_hwp_pids", return_value=["cleanup"]) as terminate,
        patch.object(hwp_com.subprocess, "run", side_effect=timeout),
    ):
        try:
            hwp_com.run_hwp_preflight(timeout=45)
        except hwp_com.HwpPreflightTimeoutError as exc:
            error = exc

    assert error is not None
    terminate.assert_called_once_with({200, 300})
