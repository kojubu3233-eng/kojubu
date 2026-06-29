import streamlit as st
import pandas as pd
from datetime import datetime

# 1. 페이지 기본 설정
st.set_page_config(page_title="종합 매출 대시보드", layout="wide")

st.title("📊 매출 분석 실시간 대시보드")
st.write("💡 월 1회 깃허브에 업로드된 엑셀 데이터를 기반으로 자동 업데이트됩니다.")

# 2. 데이터 로드 함수 (파일 업로더 없이 깃허브 파일 직접 읽기)
@st.cache_data(ttl=3600) # 1시간마다 캐시 갱신 (데이터 업데이트 반영 목적)
def load_full_data():
    try:
        # 깃허브에 올라간 엑셀 파일 이름과 똑같아야 합니다.
        df = pd.read_excel("data.xlsx")
        
        # 열 이름 공백 제거
        df.columns = df.columns.str.strip()
        
        # [날짜 전처리]
        if '일자' in df.columns:
            df['일자'] = pd.to_datetime(df['일자'], errors='coerce')
            df = df.dropna(subset=['일자'])
        else:
            st.error("⚠️ 엑셀 파일에 '일자' 열이 없습니다.")
            return None
            
        # [매출 전처리] 소수점 연산 오차 원천 차단
        if '매출' in df.columns:
            df['매출'] = pd.to_numeric(df['매출'], errors='coerce').fillna(0)
            df['매출'] = df['매출'].round(0).astype(int)
        else:
            st.error("⚠️ 엑셀 파일에 '매출' 열이 없습니다.")
            return None
            
        return df
        
    except FileNotFoundError:
        st.error("⚠️ 깃허브 저장소에 'data.xlsx' 파일이 없습니다. 파일을 업로드해 주세요.")
        return None
    except Exception as e:
        st.error(f"⚠️ 데이터 읽기 실패: {e}")
        return None

# 데이터 불러오기
df = load_full_data()

if df is not None:
    # 3. 기준 날짜 설정 (접속한 현재 시간 기준)
    current_date = datetime.now()
    current_year = current_date.year
    current_month = current_date.month
    
    if current_month == 1:
        prev_year = current_year - 1
        prev_month = 12
    else:
        prev_year = current_year
        prev_month = current_month - 1

    # 4. 핵심 매출 지표 계산
    this_year_df = df[df['일자'].dt.year == current_year]
    total_sales_year = this_year_df['매출'].sum()
    
    this_month_df = this_year_df[this_year_df['일자'].dt.month == current_month]
    total_sales_month = this_month_df['매출'].sum()
    
    prev_month_df = df[(df['일자'].dt.year == prev_year) & (df['일자'].dt.month == prev_month)]
    total_sales_prev_month = prev_month_df['매출'].sum()

    # 5. 대시보드 상단 요약 UI
    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label=f"{current_year}년 누적 매출액", value=f"{total_sales_year:,} 원")
    with col2:
        st.metric(label=f"{current_month}월 당월 매출액", value=f"{total_sales_month:,} 원")
    with col3:
        st.metric(label=f"{prev_month}월 전월 매출액", value=f"{total_sales_prev_month:,} 원")
    st.divider()

    # 6. 시각화 그래프 섹션
    graph_col1, graph_col2 = st.columns(2)
    
    with graph_col1:
        st.subheader("📈 월별 매출 추이")
        monthly_trend = this_year_df.groupby(this_year_df['일자'].dt.to_period('M'))['매출'].sum().reset_index()
        monthly_trend['일자'] = monthly_trend['일자'].astype(str)
        monthly_trend = monthly_trend.set_index('일자')
        st.bar_chart(monthly_trend)

    with graph_col2:
        st.subheader("🏆 당월 인기 품목 TOP 5")
        if '품목명' in df.columns:
            top_items = this_month_df.groupby('품목명')['매출'].sum().reset_index()
            top_items = top_items.sort_values(by='매출', ascending=False).head(5)
            top_items = top_items.set_index('품목명')
            st.bar_chart(top_items)
        else:
            st.info("품목별 순위를 보려면 '품목명' 열이 필요합니다.")

    st.divider()

    # 7. 사이드바 필터 및 전체 데이터 표
    st.sidebar.header("🔍 데이터 필터")
    if '품목명' in df.columns:
        all_items = ["전체"] + list(df['품목명'].unique())
        selected_item = st.sidebar.selectbox("품목 선택", all_items)
        if selected_item != "전체":
            df = df[df['품목명'] == selected_item]
            
    st.subheader("📋 전체 매출 데이터 조회")
    df_display = df.sort_values(by='일자', ascending=False)
    df_display['일자'] = df_display['일자'].dt.strftime('%Y-%m-%d')
    st.dataframe(df_display, use_container_width=True)
