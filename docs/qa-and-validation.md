# QA and HWPX Validation

이 문서는 실제 HWP COM 변환 QA와 HWPX 내부 구조 검증 명령을 정리한다.

## 빠른 검증

```bash
python -B -m pytest -q
python -B -m py_compile $(rg --files -g '*.py')
python -B to_hwpx_com.py --list-formats
python -B to_hwpx_com.py --preflight --startup-timeout 20
```

strict 패턴과 파일 크기 확인:

```bash
rg -n "\bAny\b|\bcast\(|type: ignore|\bobject\b" --glob "*.py" --glob "!tests/**" .
python -B -c "from pathlib import Path; [print(f'{p}: {sum(1 for line in p.read_text(encoding=\"utf-8\").splitlines() if line.strip() and not line.lstrip().startswith(\"#\"))}') for p in sorted(Path('.').rglob('*.py')) if '.git' not in p.parts and '__pycache__' not in p.parts and sum(1 for line in p.read_text(encoding='utf-8').splitlines() if line.strip() and not line.lstrip().startswith('#')) > 250]"
```

## 실제 변환 QA

실제 HWP COM 변환과 HWPX 내부 검증, 병합 셀 HWP open/save roundtrip을 한 번에 실행한다.

```bash
python -B scripts/qa_real_conversion.py --startup-timeout 20 --conversion-timeout 120 --roundtrip-timeout 90
```

기본 산출물 위치:

```text
C:\tmp\to_hwpx_real_qa\run-<id>\
├── inputs\
├── out\
│   ├── minimal.hwpx
│   ├── table.hwpx
│   └── merged.hwpx
├── merged-roundtrip.hwp
└── qa-report.txt
```

HWP open/save roundtrip을 제외하고 변환과 내부 검증만 확인하려면:

```bash
python -B scripts/qa_real_conversion.py --skip-open-roundtrip
```

`--conversion-timeout`은 각 입력 파일 변환 subprocess 제한 시간이다. timeout이 발생하면 QA runner는 timeout 직후 새로 생긴 `Hwp.exe` 자동화 프로세스 정리를 시도하고 `qa-report.txt`에 cleanup 결과를 남긴다.

실제 변환이 멈추는 지점을 좁힐 때는 단일 입력에 `--diagnose-stages`를 붙인다.

```bash
python -B to_hwpx_com.py C:\tmp\to_hwpx_real_qa\run-<id>\inputs\minimal.md -o C:\tmp\to_hwpx_real_qa\stage-out --startup-timeout 20 --diagnose-stages
```

stderr에 다음 단계가 순서대로 출력된다.

```text
[HWP-CONVERT] minimal.md: parse_source
[HWP-CONVERT] minimal.md: XHwpDocuments.Add
[HWP-CONVERT] minimal.md: build_doc
[HWP-CONVERT] minimal.md: SaveAs
[HWP-CONVERT] minimal.md: doc.Close
[HWP-CONVERT] minimal.md: postprocess
[HWP-CONVERT] minimal.md: finalize
```

## HWPX 내부 검증

생성된 `.hwpx` 파일의 ZIP 패키지, `header.xml` borderFill, `section*.xml` 표 구조를 검증한다.

```bash
python -B hwpx_validator.py path\to\document.hwpx
python -B hwpx_validator.py C:\tmp\to_hwpx_real_qa\run-<id>\out\minimal.hwpx C:\tmp\to_hwpx_real_qa\run-<id>\out\table.hwpx
```

검증 항목:

- `mimetype` 첫 ZIP entry 및 저장 방식
- 필수 `Contents/header.xml`, `Contents/section*.xml`
- `hh:borderFills itemCnt`, 중복 ID, 정의되지 않은 `borderFillIDRef`
- 표 `rowCnt`, `colCnt`, 셀 개수
- `cellAddr`, `cellSpan`, `cellSz` 필수 속성과 범위
- 헤더 셀 borderFill 음영 참조
- 헤더/본문 borderFill 분리
- 병합 없는 행의 `cellSz width` 합계와 table width 일치

## 시각 회귀 확인

`qa_real_conversion.py`가 만든 산출물을 HWP에서 열어 다음 항목을 확인한다. 자동 PDF/image export가 안정화되기 전까지는 이 목록을 수동 QA 기준으로 사용한다.

```text
C:\tmp\to_hwpx_real_qa\run-<id>\out\table.hwpx
C:\tmp\to_hwpx_real_qa\run-<id>\out\merged.hwpx
```

확인 기준:

- 표가 본문 폭을 크게 벗어나지 않는다.
- 헤더 행에 회색 음영이 적용되어 있고 본문 셀과 borderFill이 분리되어 보인다.
- 모든 셀에 기본 테두리가 보인다.
- 일정표의 `내용` 열, 정의표의 `정의` 열, 체크리스트의 `점검 항목` 열처럼 넓어야 하는 열이 가장 넓다.
- 병합 헤더 표에서 병합된 셀이 HWP에서 열고 다시 저장해도 깨지지 않는다.
- 빈 셀과 긴 본문 셀이 인접 텍스트나 다음 행을 침범하지 않는다.

HWP에서 PDF로 저장할 수 있는 환경이라면 `table.hwpx`, `merged.hwpx`를 각각 PDF로 저장하고 같은 항목을 PDF에서도 확인한다. 자동화 실패가 반복되면 PDF export 실패는 구조 검증 실패와 분리해 `qa-report.txt`에 수동 확인 필요로 기록한다.

## HWP 프로세스 확인

QA 실패 후 자동화 프로세스가 남았는지 확인한다.

```powershell
Get-CimInstance Win32_Process |
  Where-Object { $_.Name -match '^Hwp\.exe$' } |
  Select-Object ProcessId,Name,CommandLine
```

정리할 때는 다른 사용자의 열린 문서가 아닌 자동화 프로세스인지 `CommandLine`을 먼저 확인한다.
