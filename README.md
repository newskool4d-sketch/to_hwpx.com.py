# to_hwpx_com

Markdown(`.md`) 및 DOCX(`.docx`) 파일을 한글(HWP) HWPX 형식으로 변환하는 Python 스크립트입니다.  
**확장자를 자동 감지**하여 적절한 파서를 선택하며, HWP COM 자동화 방식으로 변환합니다.

## 요구 사항

- Windows OS
- [한글(HWP)](https://www.hancom.com) 설치 (COM 자동화 지원 버전)
- Python 3.8 이상
- 아래 패키지 설치

```bash
pip install python-docx pywin32
```

## 사용법

### Markdown 변환

```bash
python to_hwpx_com.py 문서.md
```

### DOCX 변환

```bash
python to_hwpx_com.py 보고서.docx
```

### 혼용 (md + docx 동시 변환)

```bash
python to_hwpx_com.py 문서.md 보고서.docx 계획.md
```

### 출력 폴더 지정

```bash
python to_hwpx_com.py 문서.md 보고서.docx -o C:\출력폴더
```

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
| 끝 표시 | — | — | 원문 보존을 위해 자동 삽입하지 않음 |
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
.md 파일  ──┐
            ├→ detect_and_parse()
.docx 파일 ─┘        │
               확장자 자동 감지
                      │
           ┌──────────┴──────────┐
    parse_markdown()       parse_docx()
           └──────────┬──────────┘
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
