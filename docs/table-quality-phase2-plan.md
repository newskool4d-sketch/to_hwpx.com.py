# 표 품질 고도화 2차 준비

## 1차 완료 범위

- `table_layout.py`로 표 역할, 명시 너비, 역할 기반 너비, 기본 셀 스타일 값을 한곳에서 계산한다.
- `blocks.table()`은 선택적으로 `table_role`, `column_widths`를 받을 수 있다.
- 범피스 `시간계획표:`는 `table_role="schedule"`, `column_widths=[14, 14, 9, 49, 14]`를 fixture로 고정했다.
- COM 렌더러는 같은 레이아웃 계산값으로 열 너비와 헤더/본문 정렬을 적용한다.
- HWPX 후처리는 같은 레이아웃 계산값으로 저장된 표의 셀 너비, 헤더 플래그, 셀 여백을 보정한다.
- COM의 `CellBorderFill`, `TablePropertyDialog` 계열 셀 스타일 액션은 실제 HWP에서 장시간 대기할 수 있어 1차 안전 범위에서 제외했다.

## 2차 핵심 목표

1. HWPX 스타일 정의 기반 헤더 음영 적용
   - `Contents/header.xml`의 borderFill 목록을 파싱한다.
   - 동일한 fill/border 정의가 있으면 재사용하고, 없으면 새 borderFill을 추가한다.
   - `Contents/section0.xml`의 헤더 셀 `borderFillIDRef`를 새 정의로 연결한다.
   - 참조 ID 충돌, namespace 보존, zip entry metadata 보존을 fixture로 검증한다.

2. 기본 테두리 품질 고도화
   - 셀별 inline 조작 대신 HWPX borderFill 정의를 공유한다.
   - body/header borderFill을 분리하되, 문서 안에 존재하지 않는 ID를 참조하지 않는다.
   - 행/열 병합, `cellSpan`, 반복 헤더가 있는 표에서 깨지지 않는지 확인한다.

3. 실제 변환 대기 원인 분리
   - 현재 환경에서는 `--preflight`는 성공하지만 실제 `SaveAs` 변환이 최소 문서에서도 120초 이상 대기했다.
   - 변환 단계별 로그를 임시로 넣어 `XHwpDocuments.Add`, `build_doc`, `SaveAs`, `doc.Close`, 후처리 중 어디서 대기하는지 분리한다.
   - 검증 명령은 timeout wrapper와 자동 PID 정리를 포함해 잔여 `Hwp.exe`를 남기지 않게 한다.

4. 역할별 표 폭 모델 확장
   - `schedule`, `budget` 외에 `definition`, `comparison`, `contacts`, `checklist` 역할을 추가한다.
   - DOCX/XLSX/HTML/PDF 파서도 명시 role이 없을 때 `table_layout.infer_table_role()` 결과를 테스트로 고정한다.
   - 페이지 폭, 단 수, section margin을 반영해 `TABLE_TOTAL_WIDTH` 고정값 의존을 줄인다.

5. fixture/시각 검증 확대
   - `bumpis_syntax.md`, 예산표, 긴 본문 표, 5열 이상 표, 빈 셀 표를 fixture corpus로 둔다.
   - HWPX unzip 검사로 `section0.xml`, `header.xml`의 너비/여백/스타일 참조를 검증한다.
   - 가능하면 HWP에서 열어 PDF 또는 이미지로 렌더링해 표 폭/음영/테두리 시각 회귀를 확인한다.

## 2차 완료 기준

- 헤더 음영과 기본 테두리가 HWPX XML 정의와 셀 참조 양쪽에서 검증된다.
- 최소 문서, 2열 표, 범피스 시간계획표 fixture가 실제 `.hwpx` 저장까지 완료된다.
- 변환 실패나 대기 시 임시 파일과 자동화 `Hwp.exe`가 남지 않는다.
- `pytest`, `py_compile`, 250 pure LOC, production strict forbidden pattern scan이 모두 통과한다.
