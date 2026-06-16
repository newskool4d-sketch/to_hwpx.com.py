# to_hwpx_com

Markdown(`.md`), TXT(`.txt`), DOCX(`.docx`), HTML(`.html`/`.htm`), CSV(`.csv`), XLSX(`.xlsx`), PDF(`.pdf`) 파일을 한글(HWP) HWPX 형식으로 변환하는 Python 스크립트입니다.  
**확장자를 자동 감지**하여 적절한 파서를 선택하며, HWP COM 자동화 방식으로 변환합니다.

## 요구 사항

- Windows OS
- [한글(HWP)](https://www.hancom.com) 설치 (COM 자동화 지원 버전)
- Python 3.10 이상
- 변환 기능에 맞는 패키지 설치

```bash
pip install pywin32 python-docx beautifulsoup4 openpyxl pdfplumber PyMuPDF pypdf
```

PDF 구조 추출(`opendataloader-pdf`)과 OCR(`kordoc-ai`)은 선택 기능입니다. 기능별 의존성은 [Dependency Matrix](docs/dependencies.md)를 참조하세요.

HWP COM 변환은 실행 중 HWP 프로그램을 띄울 수 있습니다. CLI는 COM 자동화를 사용하며, 자동화 흐름에 수동 GUI 클릭은 포함하지 않습니다.

## 사용법

### Markdown 변환

```bash
python to_hwpx_com.py 문서.md
```

### DOCX 변환

```bash
python to_hwpx_com.py 보고서.docx
```

### 지원 형식 확인

```bash
python to_hwpx_com.py --list-formats
```

### HWP COM 사전 점검

```bash
python to_hwpx_com.py --preflight
```

HWP 시작이 느리거나 멈추는 환경에서는 제한 시간을 조정할 수 있습니다.

```bash
python to_hwpx_com.py --preflight --startup-timeout 90
python to_hwpx_com.py 문서.md --startup-timeout 90
```

실제 변환 QA와 HWPX 내부 검증 명령은 [QA and HWPX Validation](docs/qa-and-validation.md)을 참조하세요.

### 문서 끝 표시 삽입

```bash
python to_hwpx_com.py 문서.md --insert-end-mark
```

### 스캔 PDF OCR 경로 지정

```bash
python to_hwpx_com.py 스캔.pdf --kordoc-home C:\kordoc-ai
```

### 혼용 (여러 형식 동시 변환)

```bash
python to_hwpx_com.py 문서.md 보고서.docx 자료.csv
```

### 출력 폴더 지정

```bash
python to_hwpx_com.py 문서.md 보고서.docx -o C:\출력폴더
```

`-o`/`--output-dir`을 생략하면 입력 파일과 같은 폴더에 저장합니다. 일반 변환도 먼저 bounded HWP COM 사전 점검을 실행하므로, HWP 시작이 제한 시간을 넘기면 변환을 시작하지 않고 실패합니다.

## 범피스 문법 MVP

범피스 분석자료의 보고서용 문법 중 아래 subset을 Markdown 입력에서 지원합니다. 범피스 실행 파일, GUI, ChatGPT, Excel, PowerPoint 기능은 구현 대상이 아닙니다.

```text
제목: 영상회의 개최 계획
소제목: 회의 개요
네모: 추진 배경
원: 참석 대상 안내
바: 세부 내용
별: 참고 사항
당구장: 일정은 변동될 수 있음
주석: 내부 검토용

표: 구분: 내용
표: A: 첫째

시간계획표:15:00:15:05:5’:인사 말씀:국장
```

매핑은 기존 block type 안에서 처리합니다. `제목:`은 1수준 제목, `소제목:`은 2수준 제목, `네모:`는 일반 단락, `원:`/`바:`/`별:`은 목록, `당구장:`/`주석:`은 인용/주석 단락, 연속 `표:`와 `시간계획표:` 줄은 표로 변환합니다.

## 지원 기능

| 기능 | Markdown | DOCX | 비고 |
|------|:--------:|:----:|------|
| 제목 (H1~H3) | ✅ `#` `##` `###` | ✅ `Heading 1~3` | 크기·굵기·여백 자동 적용 |
| 본문 단락 | ✅ | ✅ | 휴먼명조 13pt, 양쪽 정렬 |
| 8단계 항목 체계 | ✅ `1.` `가.` `1)` `가)` `(1)` `(가)` `①` `㉮` | — | 한국 행정문서 표준 |
| 글머리 기호 | ✅ `- ` `* ` → `•` | — | |
| DOCX 목록 | — | ✅ | 들여쓰기 레벨 0~7 자동 반영 |
| 표 | ✅ Markdown table | ✅ DOCX 표 | 내용 유형과 한글 시각 너비를 반영해 열 너비 배분 |
| 인용문 | ✅ `>` | ✅ Quote 스타일 | 기울임·들여쓰기 적용 |
| 코드 블록 | ✅ ` ``` ` | ✅ Code 스타일 | 맑은 고딕, 들여쓰기 적용 |
| 공문 헤더 | ✅ | ✅ | `수신:` `경유:` `제목:` 자동 감지 |
| 구분선 | ✅ `---` | ✅ Horizontal 스타일 | |
| 끝 표시 | 옵션 | 옵션 | `--insert-end-mark` 사용 시 자동 삽입 |
| frontmatter | ✅ skip | — | YAML `---` 블록 무시 |
| 이미지 | — | skip | 텍스트만 변환 |

## 저장 파일명

기존 HWPX 파일은 덮어쓰지 않습니다. 같은 이름의 파일이 있으면 번호를 붙여 저장합니다.

```text
문서.hwpx
문서 - 2.hwpx
문서 - 3.hwpx
```

## 폰트 설정

| 용도 | 한글 폰트 | 영문 폰트 |
|------|-----------|-----------|
| 본문·제목 | 휴먼명조 | Arial |
| 표·코드·목록 | 맑은 고딕 | Arial |

> 폰트가 없는 경우 HWP가 대체 폰트를 자동 적용합니다.

## 변환 흐름

```
입력 파일(.md/.txt/.docx/.html/.htm/.csv/.xlsx/.pdf)
                      │
              detect_and_parse()
               확장자 자동 감지
                      │
              블록 리스트 생성
      (h / p / li / table / bq / code / hr / official_header)
                      │
                 build_doc()
              HWP COM 자동화
                      │
                 .hwpx 파일
```

## 알려진 제한 사항

| 항목 | 내용 |
|------|------|
| 이미지 | 변환 대상에서 제외 (텍스트만 처리) |
| DOCX 병합 셀 | 병합 해제되어 동일 텍스트 중복 출력 가능 |
| DOCX 인라인 서식 | 굵기·기울임 등 run 단위 서식은 단락 전체 적용으로 단순화 |
| 복잡한 레이아웃 | 다단·텍스트박스·WordArt 등 미지원 |

## 버전 이력

| 버전 | 내용 |
|------|------|
| v1 | md_to_hwpx_com v3 + docx_to_hwpx_com v1 통합. 확장자 자동 감지 추가 |
| v1.1 | 기존 HWPX 덮어쓰기 방지, 표 열 너비 계산 안정화, 표 커서 복귀 안정화, 끝 표시 자동 삽입 제거 |

## 라이선스

MIT License — 자세한 내용은 [LICENSE](LICENSE) 파일 참조.
