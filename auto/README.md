# 자동화시스템 자동 연동 설정 (JSON 키 불필요)

## 구조
- 자동화시스템 Apps Script에 읽기 전용 `action=sales` 추가 (토큰 인증, 기존 기능은 변경 없음)
- 매출 조회용 **별도 배포**를 새로 생성 → 기존 웹앱 URL은 변경 없음
- Streamlit 앱이 해당 URL을 호출해 `데이터`(1~9월) + `DB_Master`(10월~) 확정 매출을 읽음

## Streamlit Secrets (Settings → Secrets)
```toml
SALES_API_URL = "https://script.google.com/macros/s/(새 배포 URL)/exec"
SALES_API_TOKEN = "(Apps Script의 SALES_TOKEN 값)"
```

## 동작
- 수량 = 확정량(개), 매출 = 확정매출 / 확정 전 '예정' 건 제외
- 10분 캐시, 사이드바 '지금 새로고침' 버튼으로 즉시 갱신
- Secrets 미설정·연동 실패 시 기존 `data.xlsx`로 자동 대체
