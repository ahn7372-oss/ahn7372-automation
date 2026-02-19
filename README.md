# 📊 Axel 업무 자동화 폴더

LG전자 영업마케팅 업무 자동화 파이썬 스크립트 모음

---

## 📁 파일 목록

| 파일명 | 기능 |
|--------|------|
| `excel_summary.py` | 엑셀 파일 읽기 & 요약 분석 |

---

## 🚀 사용법

### excel_summary.py
```bash
# 현재 폴더의 엑셀 파일 자동 감지
python excel_summary.py

# 특정 파일 지정
python excel_summary.py 영업실적.xlsx

# 특정 시트만 분석 (번호)
python excel_summary.py 영업실적.xlsx --sheet 1

# 특정 시트만 분석 (시트명)
python excel_summary.py 영업실적.xlsx --sheet "Q1실적"
