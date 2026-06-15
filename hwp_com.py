from __future__ import annotations

import importlib
import os
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path


class HwpComError(RuntimeError):
    pass


class HwpStartupError(HwpComError):
    pass


class MissingPywin32Error(HwpStartupError):
    pass


class HwpPreflightError(HwpComError):
    pass


class HwpPreflightTimeoutError(HwpPreflightError):
    pass


class HwpPreflightWorkerError(HwpPreflightError):
    pass


def utf8_subprocess_env() -> dict[str, str]:
    env = os.environ.copy()
    env.setdefault("PYTHONIOENCODING", "utf-8")
    env.setdefault("PYTHONUTF8", "1")
    return env


def format_hwp_startup_error(exc) -> str:
    return (
        "HWP COM 자동화 시작 실패. 확인할 항목: "
        "1) pywin32 설치, 2) Hancom Office HWP 설치 및 실행 가능 여부, "
        "3) HWPFrame.HwpObject COM 등록, 4) FilePathCheckDLL SecurityModule 등록, "
        "5) 현재 사용자 권한과 보안 프로그램 차단 여부. "
        "gen_py 캐시 문제가 의심되면 자동 삭제하지 말고 문서화된 복구 명령을 수동 실행하세요. "
        f"원인: {exc}"
    )


def startup_exception_types() -> tuple[type[BaseException], ...]:
    base_types: list[type[BaseException]] = [RuntimeError, AttributeError, OSError, TypeError]
    try:
        pywintypes = importlib.import_module("pywintypes")
    except ImportError:
        return tuple(base_types)
    com_error = getattr(pywintypes, "com_error", None)
    if isinstance(com_error, type) and issubclass(com_error, BaseException):
        base_types.append(com_error)
    return tuple(base_types)


def suppress_hwp_message_boxes(hwp) -> list[str]:
    set_message_box_mode = getattr(hwp, "SetMessageBoxMode", None)
    if not callable(set_message_box_mode):
        return ["HWP SetMessageBoxMode를 사용할 수 없어 팝업 억제를 건너뜀"]
    try:
        set_message_box_mode(0)
    except startup_exception_types() as exc:
        return [f"HWP SetMessageBoxMode(0) 실패: {exc}"]
    return []


def emit_preflight_stage(stage: str) -> None:
    print(f"[HWP-PREFLIGHT] {stage}", file=sys.stderr, flush=True)


def subprocess_timeout_text(value: str | bytes | None) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace").strip()
    return value.strip()


def running_hwp_pids() -> set[int]:
    if os.name != "nt":
        return set()
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Hwp.exe", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except OSError:
        return set()
    pids: set[int] = set()
    for line in result.stdout.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("INFO:"):
            continue
        parts = [part.strip().strip('"') for part in stripped.split(",")]
        if len(parts) < 2:
            continue
        try:
            pids.add(int(parts[1]))
        except ValueError:
            continue
    return pids


def terminate_hwp_pids(pids: set[int]) -> list[str]:
    receipts: list[str] = []
    if os.name != "nt":
        return receipts
    for pid in sorted(pids):
        try:
            result = subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=False,
            )
        except OSError as exc:
            receipts.append(f"taskkill PID {pid} 실행 실패: {exc}")
            continue
        output = (result.stdout or result.stderr or "").strip()
        receipts.append(f"taskkill PID {pid} exit={result.returncode}: {output}")
    return receipts


def create_hwp_object(
    visible: bool = True,
    startup_warnings: list[str] | None = None,
    stage_reporter: Callable[[str], None] | None = None,
):
    try:
        if stage_reporter is not None:
            stage_reporter("import win32com.client")
        win32com_client = importlib.import_module("win32com.client")
    except ImportError as exc:
        raise MissingPywin32Error("HWP COM 자동화에는 pywin32가 필요함: pip install pywin32") from exc
    startup_exceptions = startup_exception_types()
    try:
        if stage_reporter is not None:
            stage_reporter("dispatch HWPFrame.HwpObject")
        hwp = win32com_client.Dispatch("HWPFrame.HwpObject")
        if stage_reporter is not None:
            stage_reporter("suppress message boxes")
        warnings = suppress_hwp_message_boxes(hwp)
        if startup_warnings is not None:
            startup_warnings.extend(warnings)
        if stage_reporter is not None:
            stage_reporter("register FilePathCheckDLL")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        if stage_reporter is not None:
            stage_reporter("set window visibility")
        hwp.XHwpWindows.Item(0).Visible = visible
        if stage_reporter is not None:
            stage_reporter("startup complete")
        return hwp
    except startup_exceptions as exc:
        raise HwpStartupError(format_hwp_startup_error(exc)) from exc


def run_hwp_preflight_worker(visible: bool = False) -> str:
    hwp = None
    startup_warnings: list[str] = []
    try:
        hwp = create_hwp_object(
            visible=visible,
            startup_warnings=startup_warnings,
            stage_reporter=emit_preflight_stage,
        )
        for warning in startup_warnings:
            print(f"[WARN] {warning}", file=sys.stderr)
        return "HWP COM preflight OK: HWPFrame.HwpObject 생성 및 SecurityModule 등록 성공"
    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except startup_exception_types() as exc:
                print(f"[WARN] HWP 종료 실패: {exc}", file=sys.stderr)


def run_hwp_preflight(visible: bool = False, timeout: int = 45) -> str:
    before_hwp_pids = running_hwp_pids()
    cmd = [
        sys.executable,
        str(Path(__file__).with_name("to_hwpx_com.py").resolve()),
        "--_preflight-worker",
    ]
    if visible:
        cmd.append("--_preflight-visible")
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
            env=utf8_subprocess_env(),
        )
    except subprocess.TimeoutExpired as exc:
        after_hwp_pids = running_hwp_pids()
        cleanup_receipts = terminate_hwp_pids(after_hwp_pids - before_hwp_pids)
        output = subprocess_timeout_text(exc.stdout)
        error = subprocess_timeout_text(exc.stderr)
        detail = "\n".join(part for part in (output, error) if part)
        cleanup_detail = "\n".join(cleanup_receipts)
        suffix = f"\n부분 출력:\n{detail}" if detail else ""
        cleanup_suffix = f"\ncleanup:\n{cleanup_detail}" if cleanup_detail else ""
        raise HwpPreflightTimeoutError(
            f"HWP COM preflight timed out after {timeout} seconds.{suffix}{cleanup_suffix}",
        ) from exc
    output = (result.stdout or "").strip()
    error = (result.stderr or "").strip()
    if result.returncode == 0:
        return output or "HWP COM preflight OK"
    raise HwpPreflightWorkerError(error or output or "HWP COM preflight failed.")
