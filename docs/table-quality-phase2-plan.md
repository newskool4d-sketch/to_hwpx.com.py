# 표 품질 고도화 2차 진행 현황

## 완료된 범위

- `table_layout.py`, `table_roles.py`, `table_grid.py`로 표 역할 추론, 열 폭 계산, 병합 셀 정규화를 분리했다.
- `schedule`, `budget`, `definition`, `comparison`, `contacts`, `checklist` 역할별 폭 모델을 추가했다.
- DOCX/XLSX/HTML/PDF/CSV 파서 입력도 명시 role이 없으면 `table_layout_from_block()`에서 역할을 추론하도록 테스트를 고정했다.
- HWPX 후처리에서 `Contents/header.xml`의 `borderFill`을 재사용하거나 추가하고, `Contents/section0.xml`의 헤더/본문 셀 `borderFillIDRef`를 분리한다.
- 헤더 음영, 본문 테두리, 셀 여백, 병합 셀 span, 페이지 본문 폭 fallback을 HWPX unzip 테스트로 검증했다.
- 실제 변환 QA는 timeout wrapper와 새로 생성된 HWP 프로세스 정리를 포함하도록 `scripts/qa_real_conversion.py`, `scripts/qa_command.py`로 보강했다.
- `hwpx_validator.py`는 HWPX package entry, borderFill 참조, 헤더 fill, 헤더/본문 borderFill 분리, 병합 셀 범위, 행 폭 합계를 검사한다.
- 표 전용 validator 검사는 `hwpx_validator_tables.py`로 분리하고, `hwpx_validator.py`는 package 읽기, section 순회, CLI report에 집중하도록 정리했다.
- `tests/fixtures/table_corpus.json`으로 예산표, 긴 본문 표, 5열 일정표, 빈 셀 정의표, 비교표, 연락처표, 체크리스트표, 병합 헤더 예산표 fixture를 추가했다.
- `--preflight --startup-timeout 20` 정상 실행 후 `Hwp.exe` 잔류 프로세스가 없는 것을 확인했다.
- 직접 HWPX 생성 경로는 `secPr`의 `pagePr/margin`에서 본문 폭을 계산해 테이블 `hp:sz`와 셀 폭을 생성한다.
- COM 렌더러의 `insert_table()`/`build_doc()`에는 page width 주입점을 열었고, SaveAs 후처리는 원본 table block을 받아 저장된 section 폭 기준으로 layout을 다시 계산한다.
- 실제 변환 hang 진단용 `--diagnose-stages`를 추가해 `parse_source`, `XHwpDocuments.Add`, `build_doc`, `SaveAs`, `doc.Close`, `postprocess`, `finalize` 진행 지점을 stderr로 확인할 수 있게 했다.
- real conversion QA는 `qa_models.py`, `qa_samples.py`, `qa_report.py`, `qa_command.py`로 책임을 나누고 `qa_real_conversion.py`는 실행 오케스트레이션만 맡도록 분리했다.

## 남은 핵심 작업

1. 시각 회귀 검증
   - HWP 자동화가 안정적인 환경에서는 HWPX/HWP를 PDF 또는 이미지로 export한 뒤 표 폭, 헤더 음영, 테두리, 병합 셀을 확인한다.
   - 자동화가 불안정하면 먼저 수동 QA 체크리스트와 산출물 경로를 문서화한다.

2. 공통 표 로직 이식 준비
   - `hwpx-converter-renewal` 대상 저장소가 확인되면 `docs/hwpx-converter-renewal-table-migration.md` 순서대로 characterization 후 복사 이식한다.

## 다음 회차 권장 순서

1. HWP 렌더링 기반 PDF/image 회귀 검증을 별도 느린 QA로 둔다.
2. `hwpx-converter-renewal` 이식 전에는 현재 저장소의 표 모듈 계약과 validator 결과를 문서화한다.
3. QA 스크립트에 PDF/image export 자동화를 추가할 때는 별도 느린 QA 모듈로 둔다.

## 완료 기준

- `pytest -q`, `py_compile`, production strict forbidden pattern scan, 250 pure LOC 점검이 통과한다.
- 최소 문서, 2열 표, 범피스 시간계획표, fixture corpus가 HWPX 구조 검증을 통과한다.
- 변환 실패나 대기 시 임시 파일과 자동화 `Hwp.exe`가 남지 않는다.
- `hwpx-converter-renewal`로 옮길 표 공통 로직 후보와 테스트 기준이 문서로 남아 있다.
