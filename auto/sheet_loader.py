"""자동화시스템 Apps Script 웹앱(action=sales) -> 매출분석용 DataFrame 로더.

JSON 키 없이 동작합니다. Streamlit Secrets에 아래 2개만 등록하세요.
    SALES_API_URL   = "https://script.google.com/macros/s/....../exec"   # 매출 조회용 배포 URL
    SALES_API_TOKEN = "Apps Script의 SALES_TOKEN 값"

Apps Script 쪽 데이터 규칙 (salesData_)
  - 자동화시스템 '데이터'(1~9월) + 'DB_Master'(10월~) 병합, 고유ID 중복 시 뒤(DB_Master) 우선
  - 수량 = 확정량(개), 매출 = 확정매출, 확정 전 '예정' 건(확정매출 0)은 제외
"""
import pandas as pd
import requests
import streamlit as st

# 구분 표기 통일: 시트 표기 -> 기존 data.xlsx 표기
CATEGORY_MAP = {"국": "국/탕"}


def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.strip(),
                         errors="coerce")


@st.cache_data(ttl=600, show_spinner="자동화시스템에서 최신 매출 데이터를 불러오는 중...")
def load_from_sheet() -> pd.DataFrame:
    url = st.secrets["SALES_API_URL"]
    token = st.secrets["SALES_API_TOKEN"]
    r = requests.get(url, params={"action": "sales", "token": token}, timeout=60)
    r.raise_for_status()
    js = r.json()
    if js.get("error"):
        raise RuntimeError(js["error"])
    df = pd.DataFrame(js["rows"], columns=js["columns"])
    for c in ["납품일", "부대명", "구분", "품목"]:
        df[c] = df[c].astype(str).str.strip()
    df["구분"] = df["구분"].replace(CATEGORY_MAP)
    for c in ["수량", "단가(Vat별도)", "매출"]:
        df[c] = _to_num(df[c])
    return df[df["매출"].fillna(0) > 0].reset_index(drop=True)
