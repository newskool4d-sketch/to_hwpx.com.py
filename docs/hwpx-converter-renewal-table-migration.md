# hwpx-converter-renewal 표 로직 이식 계획

## 목적

`to_hwpx_com.py`에서 안정화한 표 품질 로직을 `hwpx-converter-renewal`에 옮기기 위한 작업 기준이다. 장기적으로는 두 저장소가 표 역할 추론, 폭 계산, HWPX borderFill 후처리, validator, real QA fixture를 공유해야 한다.

## 현재 저장소 기준 공통 후보

우선 복사 이식할 수 있는 모듈:

- `table_roles.py`: 표 역할 상수, 역할 추론, 역할별 폭 profile
- `table_grid.py`: 병합 셀 grid 확장과 `merged_cells` 정규화
- `table_layout.py`: block 기반 `TableLayout` 생성, explicit/role/내용 기반 폭 계산
- `table_hwpx_styles.py`: `header.xml` borderFill 재사용/추가
- `table_hwpx_postprocess.py`: `section*.xml` 셀 폭, borderFill, margin, span 후처리
- `hwpx_page.py`: `secPr/pagePr/margin` 기반 본문 폭 계산
- `hwpx_validator_core.py`, `hwpx_validator_tables.py`, `hwpx_validator.py`: HWPX 구조/표/style 검증

테스트 fixture:

- `tests/fixtures/table_corpus.json`
- `tests/test_table_fixture_corpus.py`
- `tests/test_table_roles.py`
- `tests/test_table_layout_mvp.py`
- `tests/test_table_phase2.py`
- `tests/test_table_hwpx_postprocess.py`
- `tests/test_hwpx_validator.py`

QA 도구:

- `scripts/qa_models.py`
- `scripts/qa_samples.py`
- `scripts/qa_command.py`
- `scripts/qa_report.py`
- `scripts/qa_real_conversion.py`
- `scripts/hwp_roundtrip_check.py`

## 이식 전 characterization

대상 저장소에서 먼저 현재 동작을 테스트로 고정한다.

```powershell
cd "C:\Users\홍주형\projects\hwpx-converter-renewal"
git status --short --branch
python -B -m pytest -q
python -B -m py_compile $(rg --files -g '*.py')
python -B to_hwpx_com.py --list-formats
python -B to_hwpx_com.py --preflight --startup-timeout 20
```

대상 CLI 이름이 다르면 `to_hwpx_com.py` 부분만 실제 entry point로 바꾼다.

최소 characterization 항목:

- `--help`, `--list-formats`
- unsupported extension의 exit code와 메시지
- output path numbering 또는 overwrite 정책
- preflight 실패 시 exit code
- parser가 반환하는 table block shape
- COM renderer와 direct HWPX renderer 중 어느 경로가 기본인지

## 권장 이식 순서

1. 표 block 계약부터 맞춘다.
   - `header: list[str]`
   - `rows: list[list[str]]`
   - `table_role?: str`
   - `column_widths?: list[int]`
   - `table_source?: str`
   - `merged_cells?: list[list[int]]` where each span is `[row, col, row_span, col_span]`

2. 순수 표 로직을 먼저 옮긴다.
   - `table_roles.py`
   - `table_grid.py`
   - `table_layout.py`
   - `hwpx_page.py`

3. 순수 테스트를 먼저 통과시킨다.
   - `test_table_roles.py`
   - `test_table_layout_mvp.py`
   - `test_table_phase2.py`
   - `test_table_fixture_corpus.py`의 역할/폭 계산 부분

4. HWPX XML 후처리를 옮긴다.
   - `table_hwpx_styles.py`
   - `table_hwpx_postprocess.py`
   - 대상 저장소의 HWPX package 구조에 맞춰 `Contents/header.xml`, `Contents/section*.xml` 경로를 확인한다.

5. validator를 옮긴다.
   - `hwpx_validator_core.py`
   - `hwpx_validator_tables.py`
   - `hwpx_validator.py`

6. renderer 연결을 한다.
   - COM renderer는 `insert_table(..., total_width=...)` 주입점을 둔다.
   - direct HWPX renderer는 `secPr`에서 `content_width_from_secpr()`를 호출해 표 생성 폭으로 넘긴다.
   - SaveAs 후처리는 미리 계산된 고정 `TableLayout`이 아니라 원본 table block을 넘겨 저장된 section 폭 기준으로 다시 계산한다.

7. real QA를 옮긴다.
   - `qa_real_conversion.py --skip-open-roundtrip`으로 구조 검증
   - 병합 셀은 `hwp_roundtrip_check.py`로 HWP open/save 확인
   - timeout cleanup receipt와 `Hwp.exe` 잔류 확인

## 완료 기준

- `table_corpus.json` fixture가 대상 저장소에서도 역할/폭 계산을 통과한다.
- HWPX unzip 검사에서 header/body borderFill 분리, 헤더 fillBrush, 행 폭 합계, 병합 span이 통과한다.
- 직접 HWPX 경로가 `secPr/pagePr/margin` 본문 폭으로 table width를 생성한다.
- COM SaveAs 후처리가 section 폭 기준으로 table layout을 다시 계산한다.
- 병합 셀 HWPX가 HWP에서 열리고 HWP로 다시 저장된다.
- `pytest -q`, `py_compile`, forbidden pattern scan, 250 pure LOC 점검이 통과한다.
- real QA 실패 시 새로 생긴 HWP 자동화 프로세스만 cleanup하고 receipt를 남긴다.

## 장기 공유 전략

초기에는 대상 저장소에 복사 이식한다. 두 저장소의 표 계약과 fixture가 같아진 뒤 아래 중 하나로 공통화한다.

- `hwpx_table_common/` 패키지로 분리
- 별도 internal repo/submodule
- 한 저장소를 source of truth로 두고 release artifact를 vendor

공통화 전제:

- public API는 `TableLayout`, `table_layout_from_block()`, `apply_table_width_profiles()`, `validate_hwpx()` 정도로 좁힌다.
- parser별 block shape를 fixture로 고정한다.
- HWP COM 의존 QA와 순수 HWPX XML 검증을 분리한다.
