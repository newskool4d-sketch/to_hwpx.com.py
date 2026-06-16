from __future__ import annotations

from pathlib import Path

from scripts.hwp_roundtrip_check import roundtrip_worker_command


def test_roundtrip_worker_command_uses_script_worker_entrypoint(tmp_path: Path) -> None:
    # Given
    hwpx_path = tmp_path / "merged.hwpx"
    output_path = tmp_path / "merged.hwp"

    # When
    command = roundtrip_worker_command("python", hwpx_path, output_path)

    # Then
    assert command[:3] == ["python", "-B", str(Path("scripts/hwp_roundtrip_check.py"))]
    assert command[-3:] == ["--worker", str(hwpx_path), str(output_path)]
