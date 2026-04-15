# md_to_hwpx_com

Markdown 파일을 한글(HWP) HWPX 형식으로 변환하는 Python 스크립트입니다.  
**HWP COM 자동화** 방식을 사용하며, 한국 행정문서 8단계 항목 체계를 지원합니다.

## 요구 사항

- Windows OS
- [한글(HWP)](https://www.hancom.com) 설치 (COM 자동화 지원 버전)
- Python 3.8 이상
- `pywin32` 패키지

```bash
pip install pywin32
```

## 사용법

### 기본 변환

```bash
python md_to_hwpx_com.py 문서.md
```

### 복수 파일 변환

```bash
python md_to_hwpx_com.py 문서1.md 문서2.md 문서3.md
```

### 출력 폴더 지정

```bash
python md_to_hwpx_com.py 문서.md -o C:\출력폴더
```

## 지원 기능

| 기능 | 설명 |
|------|------|
| 제목 (H1~H3) | 크기·굵기·여백 자동 적용 |
| 본문 단락 | 휴먼명조 13pt, 양쪽 정렬 |
| 표 (Markdown table) | 열 너비 한글=2·영문=1 시각 너비 비례 배분 |
| 8단계 항목 체계 | `1.` / `가.` / `1)` / `가)` / `(1)` / `(가)` / `①` / `㉮` |
| 글머리 기호 | `- ` / `* ` → `•` |
| 인용문 | `>` blockquote |
| 코드 블록 | ` ``` ` 블록 |
| 공문 헤더 | `수신:` / `경유:` / `제목:` 자동 감지 |
| 끝 표시 | 문서 마지막에 자동 삽입 |
| frontmatter | `---` YAML frontmatter 자동 skip |

## 폰트 설정

| 용도 | 한글 폰트 | 영문 폰트 |
|------|-----------|-----------|
| 본문·제목 | 휴먼명조 | Arial |
| 표·코드 | 맑은 고딕 | Arial |

> 폰트가 없는 경우 HWP가 대체 폰트를 자동 적용합니다.

## 변환 흐름

```
Markdown 파일
    ↓ parse_markdown()
블록 리스트 (h / p / li / table / bq / code / hr / official_header)
    ↓ build_doc()
HWP COM 조작 (삽입·서식·표 생성)
    ↓ SaveAs()
.hwpx 파일
```

## 버전 이력

| 버전 | 변경 내용 |
|------|-----------|
| v3 | frontmatter skip / Item(0) 참조 개선 / `-o` 출력경로 인자 추가 / 표 열 너비 시각 너비 기반 비례 배분 / 전체 너비 14000 확대 |
| v2 | 8단계 항목 체계 / 끝 표시 / 공문 구조 헤더 지원 |
| v1 | 최초 릴리스 |

## 라이선스

MIT License — 자세한 내용은 [LICENSE](LICENSE) 파일 참조.
