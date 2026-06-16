"""
HWP COM 자동화로 Markdown / TXT / DOCX / HTML / CSV / XLSX / PDF → HWPX 변환 (v2)
확장자를 자동 감지하여 적절한 파서로 처리.
이미지는 skip.

변경 이력:
  v1 - md_to_hwpx_com v3 + docx_to_hwpx_com v1 통합
  v2 - TXT / HTML / CSV / XLSX / PDF 지원 추가 (opendataloader-pdf 우선)
"""
import argparse
import os
import sys
import time
from dataclasses import dataclass
from pathlib import Path

from conversion_service import ConversionServiceError
from conversion_service import convert_file
from converter_dispatch import SUPPORTED_EXTENSIONS, detect_and_parse
from converter_dispatch import UnsupportedInputFormatError
from converter_dispatch import input_format_from_extension
from hwp_com import HwpComError
from hwp_com import create_hwp_object
from hwp_com import run_hwp_preflight
from hwp_com import run_hwp_preflight_worker as _run_hwp_preflight_worker
from hwp_com import startup_exception_types


def _configure_utf8_stdio():
    os.environ.setdefault('PYTHONIOENCODING', 'utf-8')
    for stream_name in ('stdout', 'stderr'):
        stream = getattr(sys, stream_name, None)
        if stream is not None and hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8', errors='replace')


_configure_utf8_stdio()


CONVERSION_FAILURE_TYPES: tuple[type[Exception], ...] = (ConversionServiceError, FileExistsError)


@dataclass(frozen=True, slots=True)
class OutputPathPreparationError(RuntimeError):
    source_arg: str
    original_error: OSError

    def __str__(self) -> str:
        return f'출력 경로 준비 실패: {self.original_error}'


def build_output_path(src_path, output_dir):
    return str(_select_output_path(Path(src_path), output_dir, set()))


def _output_directory(src_path: Path, output_dir: str | Path | None) -> Path:
    if output_dir:
        return Path(output_dir).expanduser().resolve()
    return src_path.expanduser().resolve().parent


def _select_output_path(src_path: Path, output_dir: str | Path | None, reserved_paths: set[Path]) -> Path:
    out_dir = _output_directory(src_path, output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    candidate = out_dir / f'{src_path.stem}.hwpx'
    if not candidate.exists() and candidate.resolve() not in reserved_paths:
        return candidate
    for idx in range(2, 1000):
        candidate = out_dir / f'{src_path.stem} - {idx}.hwpx'
        if not candidate.exists() and candidate.resolve() not in reserved_paths:
            return candidate
    raise FileExistsError(f'저장 가능한 파일명을 찾지 못함: {out_dir / f"{src_path.stem}.hwpx"}')


def positive_int(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError('양의 정수여야 함') from exc
    if parsed <= 0:
        raise argparse.ArgumentTypeError('양의 정수여야 함')
    return parsed


def validate_source_args(file_args: list[str]) -> tuple[list[tuple[str, Path]], list[tuple[str, Exception]]]:
    valid_sources: list[tuple[str, Path]] = []
    failures: list[tuple[str, Exception]] = []
    for src_arg in file_args:
        src_path = Path(src_arg).expanduser().resolve()
        if not src_path.exists():
            failures.append((src_arg, FileNotFoundError(f'입력 파일 없음: {src_path}')))
            continue
        if not src_path.is_file():
            failures.append((src_arg, IsADirectoryError(f'입력 경로가 파일이 아님: {src_path}')))
            continue
        try:
            input_format_from_extension(src_path.suffix.lower())
        except UnsupportedInputFormatError as exc:
            failures.append((src_arg, exc))
            continue
        valid_sources.append((src_arg, src_path))
    return valid_sources, failures


def plan_output_paths(
    valid_sources: list[tuple[str, Path]],
    output_dir: str | Path | None,
) -> tuple[list[tuple[str, Path, Path]], list[tuple[str, Exception]]]:
    planned_sources: list[tuple[str, Path, Path]] = []
    failures: list[tuple[str, Exception]] = []
    reserved_paths: set[Path] = set()
    for src_arg, src_path in valid_sources:
        try:
            output_path = _select_output_path(src_path, output_dir, reserved_paths)
        except OSError as exc:
            failures.append((src_arg, OutputPathPreparationError(source_arg=src_arg, original_error=exc)))
            continue
        planned_sources.append((src_arg, src_path, output_path))
        reserved_paths.add(output_path.resolve())
    return planned_sources, failures


def print_failures(failures: list[tuple[str, Exception]]) -> None:
    print('\n실패 목록:', file=sys.stderr)
    for src_arg, exc in failures:
        print(f'- {src_arg}: {exc}', file=sys.stderr)


def print_conversion_stage(stage: str) -> None:
    print(f'[HWP-CONVERT] {stage}', file=sys.stderr, flush=True)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description='Markdown / TXT / DOCX / HTML / CSV / XLSX / PDF → HWPX 변환 (HWP COM 방식)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument('files', nargs='*', help='변환할 파일 경로')
    parser.add_argument('-o', '--output-dir', default=None, help='저장할 폴더 경로 (기본: 입력 파일과 같은 폴더)')
    parser.add_argument('--insert-end-mark', action='store_true', help="문서 끝에 '끝' 표시를 자동 삽입")
    parser.add_argument('--list-formats', action='store_true', help='지원 형식 목록 출력')
    parser.add_argument('--preflight', action='store_true', help='HWP COM 실행 가능 여부만 점검하고 종료')
    parser.add_argument('--startup-timeout', type=positive_int, default=45, help='HWP COM 시작 제한 시간(초), 기본 45초')
    parser.add_argument('--kordoc-home', default=None, help='스캔 PDF OCR용 kordoc-ai 경로 (또는 KORDOC_HOME 환경변수)')
    parser.add_argument('--diagnose-stages', action='store_true', help='변환 단계별 로그를 stderr로 출력')
    parser.add_argument('--_preflight-worker', action='store_true', help=argparse.SUPPRESS)
    parser.add_argument('--_preflight-visible', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args(argv)

    if args._preflight_worker:
        try:
            print(_run_hwp_preflight_worker(visible=args._preflight_visible))
            return 0
        except HwpComError as exc:
            print(str(exc), file=sys.stderr)
            return 2

    if args.list_formats:
        print('지원 입력 형식: ' + ', '.join(sorted(SUPPORTED_EXTENSIONS)))
        return 0

    if args.preflight:
        try:
            print(run_hwp_preflight(visible=False, timeout=args.startup_timeout))
            return 0
        except HwpComError as exc:
            print(f'[FAIL] {exc}', file=sys.stderr)
            return 2

    if not args.files:
        parser.error('변환할 파일 경로가 필요함')

    hwp = None
    valid_sources, failures = validate_source_args(args.files)
    for src_arg, exc in failures:
        print(f'[FAIL] {src_arg}: {exc}', file=sys.stderr)
    if not valid_sources:
        print_failures(failures)
        return 1
    planned_sources, output_failures = plan_output_paths(valid_sources, args.output_dir)
    failures.extend(output_failures)
    for src_arg, exc in output_failures:
        print(f'[FAIL] {src_arg}: {exc}', file=sys.stderr)
    if not planned_sources:
        print_failures(failures)
        return 1
    try:
        print('HWP 실행 중...')
        try:
            run_hwp_preflight(visible=False, timeout=args.startup_timeout)
        except HwpComError as exc:
            print(f'[FAIL] HWP 시작 사전 점검 실패: {exc}', file=sys.stderr)
            return 2
        try:
            hwp = create_hwp_object(visible=True)
        except HwpComError as exc:
            print(f'[FAIL] {exc}', file=sys.stderr)
            return 2
        time.sleep(1.5)

        for src_arg, src_path, hwpx_path in planned_sources:
            try:
                print(f'변환 중: {src_path.name} → {Path(hwpx_path).name}')
                kwargs = {
                    'insert_end_mark': args.insert_end_mark,
                    'kordoc_home': args.kordoc_home,
                }
                if args.diagnose_stages:
                    kwargs['stage_reporter'] = print_conversion_stage
                convert_file(
                    hwp,
                    src_path,
                    str(hwpx_path),
                    **kwargs,
                )
            except CONVERSION_FAILURE_TYPES as exc:
                failures.append((src_arg, exc))
                print(f'[FAIL] {src_arg}: {exc}', file=sys.stderr)
    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except startup_exception_types() as exc:
                print(f'[WARN] HWP 종료 실패: {exc}', file=sys.stderr)

    if failures:
        print_failures(failures)
        return 1
    print('\n전체 변환 완료.')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
