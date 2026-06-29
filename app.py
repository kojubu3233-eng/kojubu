import streamlit as st
import pandas as pd
from datetime import datetime

# 1. 페이지 기본 설정
st.set_page_config(page_title="종합 매출 대시보드", layout="wide")

st.title("📊 매출 분석 실시간 대시보드")
st.write("💡 월 1회 깃허브에 업로드된 엑셀 데이터를 기반으로 자동 업데이트됩니다.")

# 2. 데이터 로드 함수 (파일 업로더 없이 깃허브 파일 직접 읽기)
@st.cache_data(ttl=3600) 
def load_full_data():
    try:
        # 깃허브에 올라간 엑셀 파일 읽기
        df = pd.read_excel("data.xlsx")
        
        # 열 이름 공백 제거
        df.columns = df.columns.str.strip()
        
        # [날짜 전처리] '납품일'을 찾아서 '일자'로 통일
        if '납품일' in df.columns:
            df['납품일'] = pd.to_datetime(df['납품일'], errors='coerce')
            df = df.dropna(subset=['납품일'])
            df = df.rename(columns={'납품일': '일자'})
        else:
            st.error("⚠️ 엑셀 파일에 '납품일' 열이 없습니다.")
            return None
            
        # [매출 전처리]
        if '매출' in df.columns:
            df['매출'] = pd.to_numeric(df['매출'], errors='coerce').fillna(0)
            df['매출'] = df['매출'].round(0).astype(int)
        else:
            st.error("⚠️ 엑셀 파일에 '매출' 열이 없습니다.")
            return None
            
        return df
        
    except Exception as e:
        st.error(f"⚠️ 데이터 읽기 실패: {e}")
        return None

# 데이터 불러오기
df = load_full_data()

if df is not None:
    # 3. 기준 날짜 설정
    current_date = datetime.now()
    current_year = current_date.year
    current_month = current_date.month
    
    if current_month == 1:
        prev_year, prev_month = current_year - 1, 12
    else:
        prev_year, prev_month = current_year, current_month - 1

    # 4. 매출 지표 계산
    this_year_df = df[df['일자'].dt.year == current_year]
    total_sales_year = this_year_df['매출'].sum()
    
    this_month_df = this_year_df[this_year_df['일자'].dt.month == current_month]
    total_sales_month = this_month_df['매출'].sum()
    
    prev_month_df = df[(df['일자'].dt.year == prev_year) & (df['일자'].dt.month == prev_month)]
    total_sales_prev_month = prev_month_df['매출'].sum()

    # 5. UI 출력
    st.divider()
    col1, col2, col3 = st.columns(3)
    col1.metric(label=f"{current_year}년 누적 매출액", value=f"{total_sales_year:,} 원")
    col2.metric(label=f"{current_month}월 당월 매출액", value=f"{total_sales_month:,} 원")
    col3.metric(label=f"{prev_month}월 전월 매출액", value=f"{total_sales_prev_month:,} 원")
    st.divider()

    # 6. 시각화
    graph_col1, graph_col2 = st.columns(2)
    with graph_col1:
        st.subheader("📈 월별 매출 추이")
        monthly_trend = this_year_df.groupby(this_year_df['일자'].dt.to_period('M'))['매출'].sum().reset_index()
        monthly_trend['일자'] = monthly_trend['일자'].astype(str)
        st.bar_chart(monthly_trend.set_index('일자'))

    with graph_col2:
        st.subheader("🏆 당월 인기 품목 TOP 5")
        if '품목' in df.columns:
            top_items = this_month_df.groupby('품목')['매출'].sum().sort_values(ascending=False).head(5)
            st.bar_chart(top_items)

    # 7. 상세 데이터
    st.subheader("📋 전체 매출 데이터")
    st.dataframe(df.sort_values(by='일자', ascending=False), use_container_width=True)
