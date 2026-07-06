import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import numpy as np

def safe_sorted(iterable):
    return sorted([str(x) for x in iterable if str(x) not in ('nan', 'None', '')])

st.set_page_config(page_title="연간 매출 분석 대시보드", layout="wide", initial_sidebar_state="expanded")
st.title("📊 연간 매출 분석 대시보드")
st.markdown("💡 **월 1회 깃허브에 업데이트된 엑셀 데이터**를 기반으로 중점 영업 전략을 도출하세요.")
st.markdown("---")

st.sidebar.header("📁 데이터 연동 상태")
st.sidebar.success("✅ 깃허브 data.xlsx 연동 중")

@st.cache_data(ttl=3600)
def load_data():
    try:
        return pd.read_excel("data.xlsx", engine='openpyxl')
    except Exception as e:
        st.error(f"⚠️ 'data.xlsx' 파일을 읽을 수 없습니다: {e}")
        return None

raw_df = load_data()

if raw_df is not None:
    try:
        df = raw_df.copy()

        if len(df.columns) >= 7:
            df.columns = list(df.columns[:7])
            df.columns = ['납품일', '부대명', '구분', '품목', '수량', '단가(Vat별도)', '매출']
        else:
            st.error(f"⚠️ 열 개수 부족 (현재 {len(df.columns)}개, 최소 7개 필요)")
            st.stop()

        for col in ['납품일', '부대명', '구분', '품목']:
            df[col] = df[col].astype(str).str.strip()

        df = df[pd.to_numeric(df['매출'].astype(str).str.replace(',', '', regex=False), errors='coerce').notna()]
        df = df[~df['구분'].isin(['nan', 'None', ''])]
        df = df[~df['부대명'].isin(['nan', 'None', ''])]

        for col in ['수량', '단가(Vat별도)', '매출']:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        def extract_month_safely(val):
            if pd.isna(val): return 1
            val_str = str(val).strip()
            m = re.search(r'(\d+)월', val_str)
            if m: return int(m.group(1))
            m2 = re.search(r'-(\d{1,2})-', val_str)
            if m2: return int(m2.group(1))
            try:
                ts = pd.to_datetime(val_str, errors='coerce')
                if pd.notna(ts): return ts.month
            except: pass
            return 1

        df['월'] = df['납품일'].apply(extract_month_safely).astype(int)
        df = df[df['월'].between(1, 12)].sort_values('월')

        unit_pivot_raw = df.groupby(['부대명', '월'])['매출'].sum().unstack(fill_value=0)
        unit_pivot_raw.columns = [int(c) for c in unit_pivot_raw.columns]

        unit_annual   = df.groupby('부대명')['매출'].sum()
        active_months = df.groupby('부대명')['월'].nunique()
        unit_avg      = unit_annual / active_months

        # ── 강제 1:1:1 분할 로직 (pd.qcut 활용) ──
        if len(unit_annual) >= 3:
            unit_ranks = unit_annual.rank(method='first')
            unit_grades = pd.qcut(unit_ranks, q=3, labels=['🔵 저매출', '🟡 중매출', '🔴 고매출'])
            def grade_by_rank(unit_name):
                return unit_grades.get(unit_name, '🔵 저매출')
        else:
            def grade_by_rank(unit_name):
                return '🟡 중매출'

        def fmt_m(val):
            return f"{val/1_000_000:+.1f}M" if val != 0 else "0M"

        def fmt_m_abs(val):
            return f"{val/1_000_000:.1f}M"

        tab1, tab2 = st.tabs(["📊 매출 현황 분석", "🎯 부대별 영업 전략 인사이트"])

        # ══════════════════════════════════════════
        # TAB 1
        # ══════════════════════════════════════════
        with tab1:
            st.header("1. 올해 전체 매출 현황 및 요약")
            total_sales = df['매출'].sum()
            st.metric(label="💰 올해 누적 총 매출액", value=f"{total_sales:,.0f} 원")

            monthly_total = df.groupby('월')['매출'].sum().reset_index()
            fig_total = px.line(monthly_total, x='월', y='매출', markers=True,
                title="📅 월별 전체 매출 추이",
                text=monthly_total['매출'].apply(lambda x: f"{x:,.0f}원" if x > 0 else ""))
            fig_total.update_traces(textposition="top center", line=dict(width=3, color="#1f77b4"))
            fig_total.update_xaxes(dtick=1, title="월")
            fig_total.update_yaxes(title="매출액 (원)")
            st.plotly_chart(fig_total, use_container_width=True)

            st.subheader("💡 계층형 데이터 요약 (구분별 ➔ 품목별)")
            st.caption("**[구분]** 항목을 클릭하면 품목별 월별 매출 피벗을 확인할 수 있습니다.")

            category_summary = df.groupby('구분')['매출'].sum().reset_index()
            category_summary.columns = ['구분', '총매출']
            category_summary = category_summary.sort_values('총매출', ascending=False)

            for _, cat_row in category_summary.iterrows():
                category = cat_row['구분']
                cat_total = cat_row['총매출']
                with st.expander(f"📁 {category}  |  총 매출: {cat_total:,.0f}원  (클릭하여 품목별 상세 보기)"):
                    cat_df = df[df['구분'] == category]
                    item_summary = cat_df.groupby('품목')['매출'].sum().reset_index()
                    item_summary.columns = ['품목', '총매출']
                    item_summary = item_summary.sort_values('총매출', ascending=False)

                    fig_bar = px.bar(item_summary, x='품목', y='총매출',
                        text=item_summary['총매출'].apply(lambda x: f"{x:,.0f}원"),
                        title=f"[{category}] 품목별 총 매출",
                        color='총매출', color_continuous_scale='Blues')
                    fig_bar.update_traces(textposition='outside')
                    fig_bar.update_layout(showlegend=False, coloraxis_showscale=False, height=350)
                    st.plotly_chart(fig_bar, use_container_width=True)

                    st.markdown("**📋 품목별 월별 매출 상세 (피벗)**")
                    pivot_df = pd.pivot_table(cat_df, values='매출', index=['품목'], columns=['월'],
                        aggfunc='sum', fill_value=0)
                    pivot_df.columns = [f"{int(c)}월" for c in pivot_df.columns]
                    pivot_df['총합계'] = pivot_df.sum(axis=1)
                    pivot_df = pivot_df.sort_values('총합계', ascending=False)
                    safe_pivot_df = pivot_df.copy()
                    for col in safe_pivot_df.columns:
                        safe_pivot_df[col] = safe_pivot_df[col].apply(lambda x: f"{x:,.0f}")
                    st.dataframe(safe_pivot_df, use_container_width=True)

            st.markdown("---")
            st.header("2. 매출 상세 분석 및 부대별 현황")

            st.subheader("🏢 부대별 전체 매출 현황 (마스터 리스트)")
            unit_pivot = pd.pivot_table(df, index='부대명', columns='월', values='매출', aggfunc='sum', fill_value=0)
            unit_pivot.columns = [f"{int(c)}월" for c in unit_pivot.columns]
            unit_pivot['총매출'] = unit_pivot.sum(axis=1)
            unit_pivot = unit_pivot.sort_index(ascending=True)
            safe_unit_pivot = unit_pivot.map(lambda x: f"{x:,.0f}")
            st.dataframe(safe_unit_pivot, use_container_width=True)
            st.download_button(label="📥 부대별 현황 엑셀 다운로드",
                data=unit_pivot.to_csv().encode('utf-8-sig'),
                file_name='부대별_매출_현황.csv', mime='text/csv')

            st.markdown("---")
            st.subheader("📈 상세 필터 및 전월 대비 증감율 분석")

            col1, col2, col3 = st.columns(3)
            with col1: selected_units = st.multiselect("📍 부대명 선택", options=safe_sorted(df['부대명'].unique()))
            with col2: selected_categories = st.multiselect("🏷️ 구분 선택", options=safe_sorted(df['구분'].unique()))
            with col3: selected_months = st.multiselect("📅 월 선택", options=sorted(df['월'].unique()))

            analysis_df = df.copy()
            if selected_units: analysis_df = analysis_df[analysis_df['부대명'].isin(selected_units)]
            if selected_categories: analysis_df = analysis_df[analysis_df['구분'].isin(selected_categories)]
            if selected_months: analysis_df = analysis_df[analysis_df['월'].isin(selected_months)]

            st.write(f"📊 **분석 데이터:** 총 {len(analysis_df)}건")

            if not analysis_df.empty:
                filtered_monthly = analysis_df.groupby('월')['매출'].sum().reset_index()
                all_months = pd.DataFrame({'월': range(int(df['월'].min()), int(df['월'].max()) + 1)})
                filtered_monthly = pd.merge(all_months, filtered_monthly, on='월', how='left').fillna(0)
                filtered_monthly['전월매출'] = filtered_monthly['매출'].shift(1)
                filtered_monthly['증감율(%)'] = 0.0
                has_prev = filtered_monthly['전월매출'] > 0
                filtered_monthly.loc[has_prev, '증감율(%)'] = (
                    (filtered_monthly.loc[has_prev, '매출'] - filtered_monthly.loc[has_prev, '전월매출'])
                    / filtered_monthly.loc[has_prev, '전월매출']) * 100

                title_parts = []
                if selected_units:
                    title_parts.append(" · ".join(selected_units) if len(selected_units) <= 3
                        else f"{' · '.join(selected_units[:3])} 외 {len(selected_units)-3}개")
                if selected_categories: title_parts.append(" · ".join(selected_categories))
                if selected_months: title_parts.append(f"{'/'.join(str(m)+'월' for m in sorted(selected_months))}")
                chart_title = ("📈 [ " + "  |  ".join(title_parts) + " ] 매출액 추이") if title_parts else "📈 전체 매출액 추이"

                fig_f = px.line(filtered_monthly, x='월', y='매출', markers=True, title=chart_title,
                    text=filtered_monthly['매출'].apply(lambda x: f"{x:,.0f}원" if x > 0 else ""))
                fig_f.update_traces(textposition="top center", line=dict(width=3, color="#e41a1c"))
                fig_f.update_xaxes(dtick=1)
                st.plotly_chart(fig_f, use_container_width=True)

                st.subheader("📈 월별 매출액 증감율 (MoM)")
                metrics_count = len(filtered_monthly) - 1
                if metrics_count > 0:
                    m_cols = st.columns(min(metrics_count, 6))
                    for idx in range(1, len(filtered_monthly)):
                        row = filtered_monthly.iloc[idx]
                        prev_row = filtered_monthly.iloc[idx - 1]
                        m_idx = (idx - 1) % 6
                        if m_idx == 0 and idx > 1:
                            m_cols = st.columns(min(metrics_count - (idx - 1), 6))
                        p_mon, c_mon, pct = int(prev_row['월']), int(row['월']), row['증감율(%)']
                        if prev_row['매출'] == 0 and row['매출'] == 0:
                            s_html, b_color = "<span style='color:gray;font-weight:bold;'>- 0%</span>", "gray"
                        elif prev_row['매출'] == 0 and row['매출'] > 0:
                            s_html, b_color = "<span style='color:#d9534f;font-weight:bold;'>▲ 신규 발생</span>", "#d9534f"
                        elif pct > 0:
                            s_html, b_color = f"<span style='color:#d9534f;font-weight:bold;'>▲ {pct:.1f}% 상승</span>", "#d9534f"
                        elif pct < 0:
                            s_html, b_color = f"<span style='color:#0275d8;font-weight:bold;'>▼ {abs(pct):.1f}% 하락</span>", "#0275d8"
                        else:
                            s_html, b_color = "<span style='color:gray;font-weight:bold;'>- 변동 없음</span>", "gray"
                        with m_cols[m_idx]:
                            st.markdown(f"""
                            <div style="background-color:#f8f9fa;padding:15px;border-radius:6px;
                                        border-left:5px solid {b_color};text-align:center;margin-bottom:12px;">
                                <p style="margin:0;font-size:13px;color:#6c757d;">{p_mon}월 ➔ {c_mon}월</p>
                                <p style="margin:6px 0 0 0;font-size:16px;">{s_html}</p>
                            </div>""", unsafe_allow_html=True)

                st.markdown("---")
                st.subheader("📋 전체/선택 데이터 엑셀식 상세 검색")
                with st.expander("🔍 상세 검색 및 데이터 표 펼치기", expanded=False):
                    f_col1, f_col2, f_col3, f_col4 = st.columns(4)
                    with f_col1: raw_date = st.multiselect("☑️ 납품일", options=safe_sorted(analysis_df['납품일'].unique()))
                    with f_col2: raw_unit = st.multiselect("☑️ 부대명", options=safe_sorted(analysis_df['부대명'].unique()))
                    with f_col3: raw_cat  = st.multiselect("☑️ 구분",   options=safe_sorted(analysis_df['구분'].unique()))
                    with f_col4: raw_item = st.multiselect("☑️ 품목",   options=safe_sorted(analysis_df['품목'].unique()))
                    final_table_df = analysis_df.copy()
                    if raw_date: final_table_df = final_table_df[final_table_df['납품일'].isin(raw_date)]
                    if raw_unit: final_table_df = final_table_df[final_table_df['부대명'].isin(raw_unit)]
                    if raw_cat:  final_table_df = final_table_df[final_table_df['구분'].isin(raw_cat)]
                    if raw_item: final_table_df = final_table_df[final_table_df['품목'].isin(raw_item)]
                    st.write(f"결과: **{len(final_table_df)}건**")
                    display_df = final_table_df.copy()
                    for col in ['수량', '단가(Vat별도)', '매출']:
                        display_df[col] = display_df[col].apply(lambda x: f"{x:,.0f}")
                    st.dataframe(display_df, use_container_width=True)
            else:
                st.warning("선택하신 필터 조건과 일치하는 데이터가 없습니다.")

        # ══════════════════════════════════════════
        # TAB 2
        # ══════════════════════════════════════════
        with tab2:
            st.header("🎯 부대별 영업 전략 인사이트")

            selected_base_month = st.selectbox(
                "📅 분석 기준 월 선택 (전월 대비 비교)",
                options=sorted(df['월'].unique(), reverse=True),
                index=0,
                format_func=lambda x: f"{x}월"
            )
            prev_month = selected_base_month - 1

            cur_sales  = unit_pivot_raw.get(selected_base_month, pd.Series(dtype=float))
            prev_sales = unit_pivot_raw.get(prev_month, pd.Series(dtype=float))

            unit_compare = pd.DataFrame({
                '부대명': unit_pivot_raw.index,
                '전월매출': [prev_sales.get(u, 0) if hasattr(prev_sales, 'get') else 0 for u in unit_pivot_raw.index],
                '당월매출': [cur_sales.get(u, 0)  if hasattr(cur_sales,  'get') else 0 for u in unit_pivot_raw.index],
            })
            unit_compare['증감액'] = unit_compare['당월매출'] - unit_compare['전월매출']
            unit_compare['증감율(%)'] = unit_compare.apply(
                lambda r: round((r['증감액'] / r['전월매출']) * 100, 1) if r['전월매출'] > 0 else 0, axis=1)
            unit_compare['연매출']    = unit_compare['부대명'].map(unit_annual).fillna(0)
            unit_compare['월평균매출'] = unit_compare['부대명'].map(unit_avg).fillna(0)
            unit_compare['매출등급']  = unit_compare['부대명'].map(grade_by_rank)
            
            # 당월매출 0원 삭제 코드 제거 (전체 부대 유지)
            # unit_compare = unit_compare[unit_compare['당월매출'] > 0].reset_index(drop=True)

            all_months_sorted = sorted(unit_pivot_raw.columns.tolist())
            months_up_to_base = [m for m in all_months_sorted if m <= selected_base_month]

            def calc_trend(unit_name):
                if unit_name not in unit_pivot_raw.index:
                    return '안정', 0.0
                row = unit_pivot_raw.loc[unit_name]
                recent_months = months_up_to_base[-3:]
                recent_avg = row[recent_months].mean() if recent_months else 0
                overall_avg = unit_avg.get(unit_name, 0)
                if overall_avg == 0:
                    return '안정', 0.0
                diff_pct = (recent_avg - overall_avg) / overall_avg * 100
                if diff_pct <= -20:
                    return '지속하락', diff_pct
                elif diff_pct >= 20:
                    return '상승추세', diff_pct
                else:
                    return '안정', diff_pct

            unit_compare[['추세', '추세편차(%)']] = unit_compare['부대명'].apply(
                lambda u: pd.Series(calc_trend(u))
            )

            SURGE_THRESHOLD = 30

            def surge_type(pct, prev, curr):
                if prev == 0 and curr > 0: return '신규'
                if prev == 0 and curr == 0: return '➡️ 유지' # 0원 부대가 남으면서 추가된 방어코드
                if pct >= SURGE_THRESHOLD: return '📈 급증'
                if pct <= -SURGE_THRESHOLD: return '📉 급감'
                return '➡️ 유지'

            unit_compare['변동유형'] = unit_compare.apply(
                lambda r: surge_type(r['증감율(%)'], r['전월매출'], r['당월매출']), axis=1)

            def keyword_filter(df_in, key):
                if key.strip():
                    return df_in[df_in['부대명'].str.contains(key.strip(), na=False)]
                return df_in

            COLOR_MAP = {'🔴 고매출': '#d9534f', '🟡 중매출': '#f0ad4e', '🔵 저매출': '#0275d8'}

            # ── STEP 1: 연매출 기준 등급 ──
            st.markdown("---")
            st.subheader("🏅 STEP 1 · 연매출 기준 고/중/저 매출 부대")
            st.caption("강제 3등분 분할 (연매출 기준) — 상위 33% = 🔴 고매출 / 중위 33% = 🟡 중매출 / 하위 33% = 🔵 저매출")

            col_a2, col_b2, col_c2 = st.columns([1, 2, 1])
            with col_a2: filter_grade = st.selectbox("필터", ["전체", "🔴 고매출", "🟡 중매출", "🔵 저매출"], key="grade_filter")
            with col_b2: kw_grade = st.text_input("🔍 부대명 검색", key="kw_grade", placeholder="부대명 입력...")
            with col_c2: sort_grade = st.selectbox("정렬 기준", ["매출액 내림차순", "매출액 오름차순", "부대명 가나다순", "부대명 역순"], key="sort_grade")

            grade_df = unit_compare.copy()
            if filter_grade != "전체": grade_df = grade_df[grade_df['매출등급'] == filter_grade]
            grade_df = keyword_filter(grade_df, kw_grade)
            
            # 필터에 따른 정렬 적용
            if sort_grade == "매출액 내림차순": grade_df = grade_df.sort_values('연매출', ascending=False)
            elif sort_grade == "매출액 오름차순": grade_df = grade_df.sort_values('연매출', ascending=True)
            elif sort_grade == "부대명 가나다순": grade_df = grade_df.sort_values('부대명', ascending=True)
            elif sort_grade == "부대명 역순": grade_df = grade_df.sort_values('부대명', ascending=False)

            g1, g2, g3 = st.columns(3)
            g1.metric("🔴 고매출 부대", f"{len(unit_compare[unit_compare['매출등급']=='🔴 고매출'])}개")
            g2.metric("🟡 중매출 부대", f"{len(unit_compare[unit_compare['매출등급']=='🟡 중매출'])}개")
            g3.metric("🔵 저매출 부대", f"{len(unit_compare[unit_compare['매출등급']=='🔵 저매출'])}개")

            if not grade_df.empty:
                top30 = grade_df.head(30)
                fig_grade = px.bar(top30, x='부대명', y='연매출',
                    color='매출등급',
                    color_discrete_map=COLOR_MAP,
                    title="부대별 연매출 (상위 30개 표시)",
                    text=top30['연매출'].apply(fmt_m_abs))
                fig_grade.update_traces(textposition='outside', textfont_size=13)
                fig_grade.update_layout(height=460, yaxis_tickformat=',', yaxis_title='연매출 (원)')
                fig_grade.update_xaxes(tickangle=-40)
                st.plotly_chart(fig_grade, use_container_width=True)

                disp2 = grade_df[['부대명', '매출등급', '연매출', '당월매출']].copy()
                disp2['연매출']   = disp2['연매출'].apply(lambda x: f"{x:,.0f}")
                disp2['당월매출'] = disp2['당월매출'].apply(lambda x: f"{x:,.0f}")
                st.dataframe(disp2, use_container_width=True)
            else:
                st.info("해당 조건의 부대가 없습니다.")

            # ── STEP 2: 증가 / 감소 ──
            st.markdown("---")
            st.subheader(f"📈 STEP 2 · 매출 증가 / 감소 부대 ({prev_month}월 → {selected_base_month}월)")
            st.caption("색상은 STEP 1 매출등급. 레이블은 백만(M) 단위.")

            col_a, col_b, col_c = st.columns([1, 2, 1])
            with col_a: filter_trend = st.selectbox("필터", ["전체", "증가", "감소"], key="trend_filter")
            with col_b: kw_trend = st.text_input("🔍 부대명 검색", key="kw_trend", placeholder="부대명 입력...")
            with col_c: sort_trend = st.selectbox("정렬 기준", ["증감액 내림차순", "증감액 오름차순", "부대명 가나다순", "부대명 역순"], key="sort_trend")

            trend_df = unit_compare.copy()
            if filter_trend == "증가": trend_df = trend_df[trend_df['증감액'] > 0]
            elif filter_trend == "감소": trend_df = trend_df[trend_df['증감액'] < 0]
            trend_df = keyword_filter(trend_df, kw_trend)

            # 필터에 따른 정렬 적용
            if sort_trend == "증감액 내림차순": trend_df = trend_df.sort_values('증감액', ascending=False)
            elif sort_trend == "증감액 오름차순": trend_df = trend_df.sort_values('증감액', ascending=True)
            elif sort_trend == "부대명 가나다순": trend_df = trend_df.sort_values('부대명', ascending=True)
            elif sort_trend == "부대명 역순": trend_df = trend_df.sort_values('부대명', ascending=False)

            m1, m2, m3 = st.columns(3)
            m1.metric("📈 증가 부대 수", f"{len(unit_compare[unit_compare['증감액'] > 0])}개")
            m2.metric("📉 감소 부대 수", f"{len(unit_compare[unit_compare['증감액'] < 0])}개")
            m3.metric("📊 분석 대상", f"{len(trend_df)}개")

            if not trend_df.empty:
                t30 = trend_df.head(30)
                fig_trend = px.bar(t30, x='부대명', y='증감액',
                    color='매출등급',
                    color_discrete_map=COLOR_MAP,
                    title=f"{prev_month}월 → {selected_base_month}월 부대별 매출 증감액 — 색상은 매출등급",
                    text=t30['증감액'].apply(fmt_m))
                fig_trend.update_traces(textposition='outside', textfont_size=13)
                fig_trend.update_layout(height=460, yaxis_title='증감액 (원)', yaxis_tickformat=',')
                fig_trend.update_xaxes(tickangle=-40)
                st.plotly_chart(fig_trend, use_container_width=True)

                disp = trend_df[['부대명', '매출등급', '전월매출', '당월매출', '증감액', '증감율(%)']].copy()
                for c in ['전월매출', '당월매출']:
                    disp[c] = disp[c].apply(lambda x: f"{x:,.0f}")
                disp['증감액']    = disp['증감액'].apply(lambda x: f"{x:+,.0f}")
                disp['증감율(%)'] = disp['증감율(%)'].apply(lambda x: f"{x:+.1f}%")
                st.dataframe(disp, use_container_width=True)
            else:
                st.info("해당 조건의 부대가 없습니다.")

            # ── STEP 3: 급증 / 급감 ──
            st.markdown("---")
            st.subheader(f"⚡ STEP 3 · 매출 급증 / 급감 부대 (전월 대비 ±{SURGE_THRESHOLD}% 이상)")
            st.caption(f"STEP 2 증감 중 ±{SURGE_THRESHOLD}% 이상 변동한 부대. 색상은 STEP 1 매출등급. 레이블은 백만(M) 단위.")

            col_a3, col_b3, col_c3 = st.columns([1, 2, 1])
            with col_a3: filter_surge = st.selectbox("필터", ["전체", "📈 급증", "📉 급감", "➡️ 유지", "신규"], key="surge_filter")
            with col_b3: kw_surge = st.text_input("🔍 부대명 검색", key="kw_surge", placeholder="부대명 입력...")
            with col_c3: sort_surge = st.selectbox("정렬 기준", ["증감액 내림차순", "증감액 오름차순", "부대명 가나다순", "부대명 역순"], key="sort_surge")

            surge_df = unit_compare.copy()
            if filter_surge != "전체": surge_df = surge_df[surge_df['변동유형'] == filter_surge]
            surge_df = keyword_filter(surge_df, kw_surge)

            # 필터에 따른 정렬 적용
            if sort_surge == "증감액 내림차순": surge_df = surge_df.sort_values('증감액', ascending=False)
            elif sort_surge == "증감액 오름차순": surge_df = surge_df.sort_values('증감액', ascending=True)
            elif sort_surge == "부대명 가나다순": surge_df = surge_df.sort_values('부대명', ascending=True)
            elif sort_surge == "부대명 역순": surge_df = surge_df.sort_values('부대명', ascending=False)

            s1, s2, s3 = st.columns(3)
            s1.metric("📈 급증 부대", f"{len(unit_compare[unit_compare['변동유형']=='📈 급증'])}개")
            s2.metric("📉 급감 부대", f"{len(unit_compare[unit_compare['변동유형']=='📉 급감'])}개")
            s3.metric("📌 꾸준 저매출", f"{len(unit_compare[(unit_compare['변동유형']=='➡️ 유지') & (unit_compare['매출등급']=='🔵 저매출')])}개")

            if not surge_df.empty:
                surge_chart = surge_df[surge_df['변동유형'].isin(['📈 급증', '📉 급감'])].head(30)
                if not surge_chart.empty:
                    fig_surge = px.bar(surge_chart, x='부대명', y='증감액',
                        color='매출등급',
                        color_discrete_map=COLOR_MAP,
                        title=f"급증/급감 부대 증감액 (원) — 색상은 매출등급(STEP 1 기준)",
                        text=surge_chart['증감액'].apply(fmt_m))
                    fig_surge.update_traces(textposition='outside', textfont_size=13)
                    fig_surge.update_layout(height=460, yaxis_title='증감액 (원)', yaxis_tickformat=',')
                    fig_surge.update_xaxes(tickangle=-40)
                    st.plotly_chart(fig_surge, use_container_width=True)

                disp3 = surge_df[['부대명', '매출등급', '변동유형', '추세', '전월매출', '당월매출', '증감율(%)']].copy()
                disp3['전월매출']  = disp3['전월매출'].apply(lambda x: f"{x:,.0f}")
                disp3['당월매출']  = disp3['당월매출'].apply(lambda x: f"{x:,.0f}")
                disp3['증감율(%)'] = disp3['증감율(%)'].apply(lambda x: f"{x:+.1f}%")
                st.dataframe(disp3, use_container_width=True)
            else:
                st.info("해당 조건의 부대가 없습니다.")

            # ══════════════════════════════════════════
            # STEP 4 : 영업 활동 집중 부대 추천 
            # ══════════════════════════════════════════
            st.markdown("---")
            st.subheader(f"🚀 STEP 4 · {selected_base_month + 1}월 영업 활동 집중 부대 추천")
            st.caption(
                f"1월~{selected_base_month}월 전체 추세를 종합 분석하여 4가지 패턴으로 부대를 분류합니다. "
                "단일 월의 등락율이 아닌 **전체 기간의 매출 흐름**을 근거로 판단하므로 일시적 변동에 흔들리지 않습니다."
            )

            n_months = len(months_up_to_base)

            if n_months < 3:
                st.warning("⚠️ 추세 분석을 위해서는 최소 3개월 이상의 데이터가 필요합니다. 현재 분석 가능한 월: " + f"{n_months}개월")
            else:
                def get_unit_series(unit_name):
                    if unit_name not in unit_pivot_raw.index:
                        return pd.Series([0]*n_months, index=months_up_to_base)
                    row = unit_pivot_raw.loc[unit_name]
                    return pd.Series([row.get(m, 0) for m in months_up_to_base], index=months_up_to_base)

                stats_rows = []
                for u in unit_compare['부대명']:
                    series = get_unit_series(u)
                    x = np.arange(len(series))
                    y = series.values.astype(float)

                    mean_val = y.mean()
                    std_val  = y.std()
                    cv = (std_val / mean_val) if mean_val > 0 else 0

                    if len(x) >= 2 and mean_val > 0:
                        slope, intercept = np.polyfit(x, y, 1)
                        slope_pct_of_mean = (slope / mean_val) * 100 if mean_val > 0 else 0
                    else:
                        slope, slope_pct_of_mean = 0, 0

                    if len(y) >= 4:
                        baseline_months = y[:-1] 
                        baseline_avg = baseline_months[-3:].mean() if len(baseline_months) >= 3 else baseline_months.mean()
                    else:
                        baseline_avg = y[:-1].mean() if len(y) > 1 else 0
                        
                    current_val = y[-1]
                    drop_from_baseline_pct = ((current_val - baseline_avg) / baseline_avg * 100) if baseline_avg > 0 else 0

                    stats_rows.append({
                        '부대명': u,
                        '평균매출': mean_val,
                        '표준편차': std_val,
                        '변동계수': cv,
                        '추세기울기_pct': slope_pct_of_mean,
                        '직전평균': baseline_avg,
                        '기준월매출': current_val,
                        '기준월괴리율': drop_from_baseline_pct,
                    })

                stats_df = pd.DataFrame(stats_rows)
                stats_df = stats_df.merge(unit_compare[['부대명', '매출등급', '연매출']], on='부대명', how='left')
                stats_df = stats_df[stats_df['평균매출'] > 0].reset_index(drop=True)

                cv_75 = stats_df['변동계수'].quantile(0.75)
                avg_sales_33 = stats_df['평균매출'].quantile(0.33)

                pattern1 = stats_df[stats_df['추세기울기_pct'] <= -5].copy()
                pattern1 = pattern1.sort_values('추세기울기_pct')

                pattern2 = stats_df[
                    (stats_df['변동계수'] >= cv_75) &
                    (stats_df['평균매출'] >= stats_df['평균매출'].quantile(0.10))
                ].copy()
                pattern2 = pattern2.sort_values('변동계수', ascending=False)

                pattern3 = stats_df[
                    (stats_df['변동계수'] < cv_75) &      
                    (stats_df['기준월괴리율'] <= -30) &     
                    (stats_df['직전평균'] > 0)
                ].copy()
                pattern3 = pattern3.sort_values('기준월괴리율')

                pattern4 = stats_df[
                    (stats_df['평균매출'] <= avg_sales_33) &
                    (stats_df['변동계수'] < cv_75) &
                    (stats_df['기준월매출'] > 0)  
                ].copy()
                pattern4 = pattern4.sort_values('평균매출', ascending=False)

                all_col_months = months_up_to_base
                all_col_names = [f"{m}월매출" for m in all_col_months]

                def attach_all_months(d):
                    d = d.copy()
                    for m, col in zip(all_col_months, all_col_names):
                        d[col] = d['부대명'].apply(
                            lambda u: unit_pivot_raw.loc[u, m] if u in unit_pivot_raw.index and m in unit_pivot_raw.columns else 0
                        )
                    return d

                pattern1 = attach_all_months(pattern1)
                pattern2 = attach_all_months(pattern2)
                pattern3 = attach_all_months(pattern3)
                pattern4 = attach_all_months(pattern4)

                def render_pattern_block(title, color, desc, pattern_df, metric_col=None, metric_label=None, metric_fmt=None):
                    st.markdown(f"""
                    <div style="border-left:5px solid {color};padding:12px 16px;margin-bottom:4px;
                                background-color:rgba(255,255,255,0.03);border-radius:4px;">
                        <strong style="color:{color};font-size:15px;">{title}</strong>
                        &nbsp;&nbsp;<span style="color:#aaa;font-size:13px;">({len(pattern_df)}개 부대)</span>
                        <p style="margin:8px 0 0 0;font-size:13px;color:#ccc;line-height:1.6;">
                            💡 <b>판단 기준:</b> {desc}
                        </p>
                    </div>""", unsafe_allow_html=True)

                    if pattern_df.empty:
                        st.info("해당 조건에 부합하는 부대가 없습니다.")
                        st.markdown("<br>", unsafe_allow_html=True)
                        return

                    if metric_col is None:
                        disp_cols = ['부대명', '매출등급'] + all_col_names
                        disp = pattern_df[disp_cols].head(15).copy()
                        for col in all_col_names:
                            disp[col] = disp[col].apply(lambda x: f"{x:,.0f}")
                        st.dataframe(disp, use_container_width=True)
                        st.markdown("<br>", unsafe_allow_html=True)
                        return

                    disp_cols = ['부대명', '매출등급'] + all_col_names + [metric_col]
                    disp = pattern_df[disp_cols].head(15).copy()

                    metric_formatted = disp[metric_col].apply(metric_fmt)

                    for col in all_col_names:
                        disp[col] = disp[col].apply(lambda x: f"{x:,.0f}")

                    disp[metric_col] = metric_formatted
                    disp = disp.rename(columns={metric_col: metric_label})
                    st.dataframe(disp, use_container_width=True)
                    st.markdown("<br>", unsafe_allow_html=True)

                render_pattern_block(
                    "📉 지속 하락 부대",
                    "#d9534f",
                    f"1월~{selected_base_month}월 전체 매출 추세선(회귀분석)이 일관되게 하락하는 부대입니다. 단발성 감소가 아닌 구조적 이탈 신호이므로 우선 방문하여 원인을 파악하세요.",
                    pattern1
                )

                render_pattern_block(
                    "🌊 매출 등락폭이 큰 부대",
                    "#f0ad4e",
                    f"월별 매출 변동성(변동계수)이 상위 25%에 해당하는 부대입니다. 발주가 불규칙하므로 정기 발주 전환을 제안하거나, 변동 원인(계절성/행사성 수요 등)을 파악해 대응 전략을 세우세요.",
                    pattern2
                )

                render_pattern_block(
                    "⚠️ 평소엔 안정적이었는데 이번 달 급감한 부대",
                    "#e8853d",
                    f"최근 3개월 평균은 안정적이었으나 {selected_base_month}월에만 직전 평균 대비 30% 이상 급감했습니다. 일회성 이슈(재고/예산/담당자 변경)일 가능성이 높으므로 빠르게 확인이 필요합니다.",
                    pattern3, '기준월괴리율', f'{selected_base_month}월 괴리율', lambda x: f"{x:+.1f}%"
                )

                render_pattern_block(
                    "📌 꾸준히 저매출을 유지 중인 부대",
                    "#5bc0de",
                    "월별 매출 변동 없이 꾸준히 낮은 매출 수준을 유지하고 있는 부대입니다. 거래는 안정적이나 규모가 작으므로, 품목 다양화 제안이나 추가 수요 발굴을 통해 매출 확대 가능성을 점검하세요.",
                    pattern4
                )

                st.markdown("---")
                p1d = pattern1.copy(); p1d['분류'] = '지속 하락'
                p2d = pattern2.copy(); p2d['분류'] = '등락폭 큼'
                p3d = pattern3.copy(); p3d['분류'] = '이번달 급감'
                p4d = pattern4.copy(); p4d['분류'] = '꾸준 저매출'

                dl_base_cols = ['부대명', '매출등급', '평균매출'] + all_col_names + ['분류']
                dl_all = pd.concat([
                    p1d[dl_base_cols] if not p1d.empty else pd.DataFrame(columns=dl_base_cols),
                    p2d[dl_base_cols] if not p2d.empty else pd.DataFrame(columns=dl_base_cols),
                    p3d[dl_base_cols] if not p3d.empty else pd.DataFrame(columns=dl_base_cols),
                    p4d[dl_base_cols] if not p4d.empty else pd.DataFrame(columns=dl_base_cols),
                ])

                if not dl_all.empty:
                    st.download_button(
                        label="📥 영업 집중 부대 추천 전체 다운로드",
                        data=dl_all.to_csv(index=False).encode('utf-8-sig'),
                        file_name=f'{selected_base_month+1}월_영업집중부대_추천.csv',
                        mime='text/csv'
                    )
                else:
                    st.info("추천 대상 부대가 없습니다.")

    except Exception as e:
        import traceback
        st.error(f"🚨 오류 발생: {e}")
        st.code(traceback.format_exc())
