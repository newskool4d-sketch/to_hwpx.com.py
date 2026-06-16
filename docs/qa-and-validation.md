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

## HWP 프로세스 확인

QA 실패 후 자동화 프로세스가 남았는지 확인한다.

```powershell
Get-CimInstance Win32_Process |
  Where-Object { $_.Name -match '^Hwp\.exe$' } |
  Select-Object ProcessId,Name,CommandLine
```

정리할 때는 다른 사용자의 열린 문서가 아닌 자동화 프로세스인지 `CommandLine`을 먼저 확인한다.
