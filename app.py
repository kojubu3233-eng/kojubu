import streamlit as st
import pandas as pd
from datetime import datetime

# 1. 페이지 기본 설정 (가로로 넓게)
st.set_page_config(page_title="종합 매출 대시보드", layout="wide")

# 2. 데이터 로드 및 초강력 전처리 함수
@st.cache_data(ttl=3600)
def load_full_data():
    try:
        # 엑셀 읽어오기
        df = pd.read_excel("data.xlsx", engine='openpyxl')
        df.columns = df.columns.str.strip() # 헤더 공백 제거
        
        # [핵심 1] 텍스트 날짜("01월 01일")를 진짜 날짜로 변환
        if '납품일' in df.columns:
            df = df.rename(columns={'납품일': '일자'})
            
            def parse_custom_date(x):
                if pd.isnull(x): return pd.NaT
                if isinstance(x, datetime): return x
                x_str = str(x).strip()
                # "07월 27일" 형태의 텍스트인 경우 강제로 날짜 조립
                if '월' in x_str and '일' in x_str:
                    current_yr = datetime.now().year # 현재 시스템 연도 사용
                    clean_str = x_str.replace('월', '-').replace('일', '').replace(' ', '')
                    try:
                        return pd.to_datetime(f"{current_yr}-{clean_str}")
                    except:
                        pass
                return pd.to_datetime(x, errors='coerce')
                
            df['일자'] = df['일자'].apply(parse_custom_date)
            df = df.dropna(subset=['일자']) # 날짜 변환 실패한 빈 줄 제거
        else:
            st.error("⚠️ 엑셀 파일에 '납품일' 열이 없습니다.")
            return None
            
        # [핵심 2] 매출액에 쉼표(,)가 있어도 숫자로 완벽히 변환
        if '매출' in df.columns:
            df['매출'] = df['매출'].astype(str).str.replace(',', '', regex=False)
            df['매출'] = pd.to_numeric(df['매출'], errors='coerce').fillna(0).astype(int)
        else:
            st.error("⚠️ 엑셀 파일에 '매출' 열이 없습니다.")
            return None
            
        return df
        
    except Exception as e:
        st.error(f"⚠️ 엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
        return None

# 데이터 불러오기 실행
df = load_full_data()

if df is not None and not df.empty:
    
    # 3. 기준 날짜 설정 (접속자 PC 시간이 아닌, '엑셀 데이터 중 가장 최신 날짜' 기준)
    max_date = df['일자'].max()
    current_year = max_date.year
    current_month = max_date.month
    
    if current_month == 1:
        prev_year, prev_month = current_year - 1, 12
    else:
        prev_year, prev_month = current_year, current_month - 1

    # 4. 매출 지표 계산 로직
    this_year_df = df[df['일자'].dt.year == current_year]
    total_sales_year = this_year_df['매출'].sum()
    
    this_month_df = this_year_df[this_year_df['일자'].dt.month == current_month]
    total_sales_month = this_month_df['매출'].sum()
    
    prev_month_df = df[(df['일자'].dt.year == prev_year) & (df['일자'].dt.month == prev_month)]
    total_sales_prev_month = prev_month_df['매출'].sum()

    # ==========================================
    # 5. UI 대시보드 화면 그리기 (기존 디자인 복구)
    # ==========================================
    st.title("📊 매출 분석 실시간 대시보드")
    st.write("💡 월 1회 깃허브에 업로드된 엑셀 데이터를 기반으로 자동 업데이트됩니다.")
    st.divider()

    # 상단 메인 타이틀 (캡처해주신 화면 반영)
    st.markdown("## 1. 올해 전체 매출 현황 및 요약")
    st.markdown(f"#### 💰 올해 누적 총 매출액\n# {total_sales_year:,.0f} 원")
    st.write("")
    st.write("")

    # 당월 / 전월 서브 미트릭
    col1, col2, col3 = st.columns(3)
    col1.metric(label=f"{current_year}년 누적 총 매출", value=f"{total_sales_year:,.0f} 원")
    col2.metric(label=f"{current_month}월 당월 매출액", value=f"{total_sales_month:,.0f} 원")
    col3.metric(label=f"{prev_month}월 전월 매출액", value=f"{total_sales_prev_month:,.0f} 원")
    st.divider()

    # 6. 시각화 (차트)
    st.markdown("## 2. 세부 지표 분석")
    graph_col1, graph_col2 = st.columns(2)
    
    with graph_col1:
        st.subheader("📈 월별 매출 추이")
        # 월별 합계 계산 후 차트 생성
        monthly_trend = this_year_df.groupby(this_year_df['일자'].dt.to_period('M'))['매출'].sum().reset_index()
        monthly_trend['일자'] = monthly_trend['일자'].astype(str)
        st.bar_chart(monthly_trend.set_index('일자'))

    with graph_col2:
        st.subheader("🏆 당월 인기 품목 TOP 5")
        if '품목' in df.columns:
            top_items = this_month_df.groupby('품목')['매출'].sum().sort_values(ascending=False).head(5)
            st.bar_chart(top_items)
        else:
            st.info("엑셀에 '품목' 열이 없어 차트를 표시할 수 없습니다.")

    st.divider()

    # 7. 사이드바 및 원본 데이터 표
    st.sidebar.header("🔍 데이터 필터")
    
    # 부대명 필터 추가 (캡처 화면 참고)
    if '부대명' in df.columns:
        all_units = ["전체"] + list(df['부대명'].dropna().unique())
        selected_unit = st.sidebar.selectbox("부대 선택", all_units)
        if selected_unit != "전체":
            df = df[df['부대명'] == selected_unit]

    if '품목' in df.columns:
        all_items = ["전체"] + list(df['품목'].dropna().unique())
        selected_item = st.sidebar.selectbox("품목 선택", all_items)
        if selected_item != "전체":
            df = df[df['품목'] == selected_item]

    st.markdown("## 3. 전체 매출 데이터 상세 조회")
    # 날짜 기준으로 내림차순 정렬하여 보여주기
    display_df = df.sort_values(by='일자', ascending=False)
    
    # 표에서 날짜 보기 좋게 다듬기
    display_df['일자'] = display_df['일자'].dt.strftime('%Y-%m-%d')
    st.dataframe(display_df, use_container_width=True, height=400)

else:
    st.warning("데이터가 비어있거나 날짜/매출 형식을 분석할 수 없습니다. 엑셀 파일을 다시 확인해 주세요.")
