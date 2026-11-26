"""자동화시스템 Apps Script 웹앱(action=sales) -> 매출분석용 DataFrame 로더.

JSON 키 없이 동작합니다. Streamlit Secrets에 아래 2개만 등록하세요.
    SALES_API_URL   = "https://script.google.com/macros/s/....../exec"   # 매출 조회용 배포 URL
    SALES_API_TOKEN = "Apps Script의 SALES_TOKEN 값"

Apps Script 쪽 데이터 규칙 (salesData_)
  - 자동화시스템 '데이터'(1~9월) + 'DB_Master'(10월~) 병합, 고유ID 중복 시 뒤(DB_Master) 우선
  - 확정 건: 수량 = 확정량(개), 매출 = 확정매출
  - 마감된 월(closed_through_month 이하): 확정 건만 집계
  - 미마감 월: 확정된 건은 확정매출, 확정되지 않은 건은 예정매출로 계산 (확정 건의 예정금액은 제외)
"""
import pandas as pd
import requests
import streamlit as st

from auto.units import KST, closed_through_month, month_of, year_of
from datetime import datetime

# 구분 표기 통일: 시트 표기 -> 기존 data.xlsx 표기
CATEGORY_MAP = {"국": "국/탕"}


def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.strip(),
                         errors="coerce")


def detect_closed_month(df: pd.DataFrame, tol: float = 0.05) -> int:
    """시트 데이터로 마감 월 자동 판단.
    월별로 '아직 확정되지 않은 예정 금액' 비중을 보고, 최근 월부터 거꾸로 보며 그 비중이 5% 이하인 첫 달까지를 마감으로 본다.
    (예: 11월은 모두 확정, 12월은 대부분 예정 → 11월까지 마감)
    """
    m = df["납품일"].map(month_of)
    conf = df["매출"].fillna(0)
    plan_open = df["예정매출"].where(conf <= 0, 0)
    g = pd.DataFrame({"m": m, "c": conf, "p": plan_open})
    g = g[g["m"].between(1, 12)].groupby("m")[["c", "p"]].sum()
    if g.empty:
        return closed_through_month()
    share = g["p"] / (g["c"] + g["p"]).replace(0, 1)
    # 최근 월부터 거꾸로 보며, 예정 비중이 낮은(=확정 완료) 첫 달을 마감 월로 본다.
    # (과거 월에 남은 미확정 예정 행 때문에 마감 월이 앞당겨지지 않도록)
    for mm in sorted(share.index, reverse=True):
        if share[mm] <= tol:
            return int(mm)
    return int(share.index.min()) - 1


@st.cache_data(ttl=600, show_spinner="자동화시스템에서 최신 매출 데이터를 불러오는 중...")
def load_from_sheet():
    """(매출 DataFrame, 마감 월) 반환."""
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

    # 올해 데이터만 (연도가 명시된 다른 연도 건은 월이 겹쳐 합산되므로 제외)
    this_year = datetime.now(KST).year
    yr = df["납품일"].map(year_of)
    df = df[yr.isna() | (yr == this_year)].copy()

    if "예정매출" in df.columns:   # 확정/예정 병행 형식
        df["예정매출"] = _to_num(df["예정매출"]).fillna(0)
        df["예정수량"] = _to_num(df["예정수량"]).fillna(0)
        try:
            override = st.secrets.get("CLOSED_THROUGH_MONTH")
        except Exception:
            override = None
        closed = closed_through_month(override) if override not in (None, "") else detect_closed_month(df)
        confirmed = df["매출"].fillna(0) > 0
        use_plan = (df["납품일"].map(month_of) > closed) & ~confirmed & (df["예정매출"] > 0)
        df.loc[use_plan, "매출"] = df.loc[use_plan, "예정매출"]
        df.loc[use_plan, "수량"] = df.loc[use_plan, "예정수량"]
        df = df[confirmed | use_plan]
        df = df[["납품일", "부대명", "구분", "품목", "수량", "단가(Vat별도)", "매출"]]
    else:
        closed = closed_through_month()
    return df[df["매출"].fillna(0) > 0].reset_index(drop=True), closed
