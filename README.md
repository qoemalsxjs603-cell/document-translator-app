# 문서 번역기 Gemini

PPTX, DOCX, XLSX 파일을 업로드하면 중국어 또는 영어로 번역한 파일을 생성하는 Streamlit 웹앱입니다.

## GitHub에 업로드할 파일

- app.py
- requirements.txt
- README.md

## Streamlit Secrets 설정

Streamlit Cloud의 Settings > Secrets에 아래처럼 입력하세요.

```toml
GEMINI_API_KEY = "본인의_Gemini_API_키"
GEMINI_MODEL = "gemini-2.5-flash"
```

API 키는 GitHub에 올리면 안 됩니다.

## 지원 기능

- PPTX / DOCX / XLSX 업로드
- 중국어 / 영어 번역
- 원본과 같은 파일 형식으로 다운로드
- 여러 파일 동시 업로드
- 번역 로그 엑셀 다운로드
- 세션 내 번역 캐시
- 선택형 용어집 업로드
- PPT 표 및 일부 그룹 오브젝트 처리
- Word 본문 / 표 / 헤더 / 푸터 처리
- Excel 텍스트 셀 처리

## 한계

- 이미지로 들어간 글자는 번역하지 못합니다.
- 일부 SmartArt / WordArt는 누락될 수 있습니다.
- PPT는 번역 후 글자 넘침과 폰트를 최종 확인하는 것을 권장합니다.
- 암호화된 파일은 지원하지 않습니다.