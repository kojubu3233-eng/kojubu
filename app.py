import streamlit as st
import pandas as pd
import plotly.express as px
import re

# ─────────────────────────────────────────────
# 헬퍼 함수
# ─────────────────────────────────────────────
def safe_sorted(iterable):
    return sorted([str(x) for x in iterable if str(x) not in ('nan', 'None', '')])

# ─────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────
st.set_page_config(page_title="연간 매출 분석 대시보드", layout="wide", initial_sidebar_state="expanded")
st.title("📊 연간 매출 분석 대시보드")
st.markdown("💡 **월 1회 깃허브에 업데이트된 엑셀 데이터**를 기반으로 중점 영업 전략을 도출하세요.")
st.markdown("---")

st.sidebar.header("📁 데이터 연동 상태")
st.sidebar.success("✅ 깃허브 data.xlsx 연동 중")

# ─────────────────────────────────────────────
# 데이터 로드
# ─────────────────────────────────────────────
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

        # ── 컬럼 정규화 ──
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

        # ── 부대별 월별 매출 피벗 (탭2 분석용) ──
        unit_pivot_raw = df.groupby(['부대명', '월'])['매출'].sum().unstack(fill_value=0)
        unit_pivot_raw.columns = [int(c) for c in unit_pivot_raw.columns]

        # ══════════════════════════════════════════
        # 탭 구성
        # ══════════════════════════════════════════
        tab1, tab2 = st.tabs(["📊 매출 현황 분석", "🎯 부대별 영업 전략 인사이트"])

        # ══════════════════════════════════════════
        # TAB 1 : 매출 현황 분석
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
            with col1:
                selected_units = st.multiselect("📍 부대명 선택", options=safe_sorted(df['부대명'].unique()))
            with col2:
                selected_categories = st.multiselect("🏷️ 구분 선택", options=safe_sorted(df['구분'].unique()))
            with col3:
                selected_months = st.multiselect("📅 월 선택", options=sorted(df['월'].unique()))

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
                st.caption("다중 필터로 표의 데이터를 자유롭게 걸러내세요.")
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
        # TAB 2 : 부대별 영업 전략 인사이트
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

            # 부대별 연 매출 합계
            unit_annual = df.groupby('부대명')['매출'].sum()

            unit_compare = pd.DataFrame({
                '부대명': unit_pivot_raw.index,
                '전월매출': [prev_sales.get(u, 0) if hasattr(prev_sales, 'get') else 0 for u in unit_pivot_raw.index],
                '당월매출': [cur_sales.get(u, 0)  if hasattr(cur_sales,  'get') else 0 for u in unit_pivot_raw.index],
            })
            unit_compare['증감액'] = unit_compare['당월매출'] - unit_compare['전월매출']
            unit_compare['증감율(%)'] = unit_compare.apply(
                lambda r: round((r['증감액'] / r['전월매출']) * 100, 1) if r['전월매출'] > 0 else 0, axis=1)
            unit_compare['연매출'] = unit_compare['부대명'].map(unit_annual).fillna(0)
            unit_compare = unit_compare[unit_compare['당월매출'] > 0].reset_index(drop=True)

            # 연평균 매출 등급 (연 매출 기준)
            unit_avg = df.groupby('부대명')['매출'].sum() / df['월'].nunique()
            top30 = unit_avg.quantile(0.70)
            bot30 = unit_avg.quantile(0.30)

            def grade(avg):
                if avg >= top30: return '🔴 고매출'
                elif avg >= bot30: return '🟡 중매출'
                else: return '🔵 저매출'

            unit_compare['연평균매출'] = unit_compare['부대명'].map(unit_avg)
            unit_compare['매출등급'] = unit_compare['연평균매출'].apply(grade)

            # ★ 급증/급감 기준 30%로 상향
            SURGE_THRESHOLD = 30

            def surge_type(pct, prev):
                if prev == 0: return '신규'
                if pct >= SURGE_THRESHOLD: return f'📈 급증'
                if pct <= -SURGE_THRESHOLD: return f'📉 급감'
                return '➡️ 유지'

            unit_compare['변동유형'] = unit_compare.apply(
                lambda r: surge_type(r['증감율(%)'], r['전월매출']), axis=1)

            def keyword_filter(df_in, key):
                if key.strip():
                    return df_in[df_in['부대명'].str.contains(key.strip(), na=False)]
                return df_in

            # ══════════════════════════════════════════
            # 섹션 1: 연평균 기준 고/중/저 매출 부대 (최상단)
            # ══════════════════════════════════════════
            st.markdown("---")
            st.subheader("🏅 STEP 1 · 연평균 기준 고/중/저 매출 부대")
            st.caption("연평균 매출 기준 상위 30% = 🔴 고매출 / 중위 40% = 🟡 중매출 / 하위 30% = 🔵 저매출 — 이 등급이 아래 모든 분석의 기준이 됩니다.")

            col_a2, col_b2 = st.columns([1, 2])
            with col_a2: filter_grade = st.selectbox("필터", ["전체", "🔴 고매출", "🟡 중매출", "🔵 저매출"], key="grade_filter")
            with col_b2: kw_grade = st.text_input("🔍 부대명 검색", key="kw_grade", placeholder="부대명 입력...")

            grade_df = unit_compare.copy()
            if filter_grade != "전체": grade_df = grade_df[grade_df['매출등급'] == filter_grade]
            grade_df = keyword_filter(grade_df, kw_grade).sort_values('연평균매출', ascending=False)

            g1, g2, g3 = st.columns(3)
            g1.metric("🔴 고매출 부대", f"{len(unit_compare[unit_compare['매출등급']=='🔴 고매출'])}개")
            g2.metric("🟡 중매출 부대", f"{len(unit_compare[unit_compare['매출등급']=='🟡 중매출'])}개")
            g3.metric("🔵 저매출 부대", f"{len(unit_compare[unit_compare['매출등급']=='🔵 저매출'])}개")

            if not grade_df.empty:
                fig_grade = px.bar(grade_df.head(30), x='부대명', y='연평균매출',
                    color='매출등급',
                    color_discrete_map={'🔴 고매출': '#d9534f', '🟡 중매출': '#f0ad4e', '🔵 저매출': '#0275d8'},
                    title="부대별 연평균 매출 (상위 30개)",
                    text=grade_df.head(30)['연평균매출'].apply(lambda x: f"{x:,.0f}원"))
                fig_grade.update_traces(textposition='outside')
                fig_grade.update_layout(height=420)
                fig_grade.update_xaxes(tickangle=-40)
                st.plotly_chart(fig_grade, use_container_width=True)

                disp2 = grade_df[['부대명', '매출등급', '연평균매출', '당월매출']].copy()
                disp2['연평균매출'] = disp2['연평균매출'].apply(lambda x: f"{x:,.0f}")
                disp2['당월매출']   = disp2['당월매출'].apply(lambda x: f"{x:,.0f}")
                st.dataframe(disp2, use_container_width=True)
            else:
                st.info("해당 조건의 부대가 없습니다.")

            # ══════════════════════════════════════════
            # 섹션 2: 매출 증가 / 감소 부대
            # ══════════════════════════════════════════
            st.markdown("---")
            st.subheader(f"📈 STEP 2 · 매출 증가 / 감소 부대 ({prev_month}월 → {selected_base_month}월)")
            st.caption("STEP 1의 매출 등급(색상)과 함께 확인하면 어느 등급 부대가 늘고 줄었는지 파악할 수 있습니다.")

            col_a, col_b = st.columns([1, 2])
            with col_a: filter_trend = st.selectbox("필터", ["전체", "증가", "감소"], key="trend_filter")
            with col_b: kw_trend = st.text_input("🔍 부대명 검색", key="kw_trend", placeholder="부대명 입력...")

            trend_df = unit_compare.copy()
            if filter_trend == "증가": trend_df = trend_df[trend_df['증감액'] > 0]
            elif filter_trend == "감소": trend_df = trend_df[trend_df['증감액'] < 0]
            trend_df = keyword_filter(trend_df, kw_trend).sort_values('증감율(%)', ascending=False)

            m1, m2, m3 = st.columns(3)
            m1.metric("📈 증가 부대 수", f"{len(trend_df[trend_df['증감액'] > 0])}개")
            m2.metric("📉 감소 부대 수", f"{len(trend_df[trend_df['증감액'] < 0])}개")
            m3.metric("📊 분석 대상", f"{len(trend_df)}개")

            if not trend_df.empty:
                # 매출등급을 색상으로 반영
                fig_trend = px.bar(trend_df.head(30), x='부대명', y='증감율(%)',
                    color='매출등급',
                    color_discrete_map={'🔴 고매출': '#d9534f', '🟡 중매출': '#f0ad4e', '🔵 저매출': '#0275d8'},
                    title=f"{prev_month}월 → {selected_base_month}월 부대별 매출 증감율 (상위 30개) — 색상은 매출등급",
                    text=trend_df.head(30)['증감율(%)'].apply(lambda x: f"{x:+.1f}%"))
                fig_trend.update_traces(textposition='outside')
                fig_trend.update_layout(height=420)
                fig_trend.update_xaxes(tickangle=-40)
                st.plotly_chart(fig_trend, use_container_width=True)

                disp = trend_df[['부대명', '매출등급', '전월매출', '당월매출', '증감액', '증감율(%)']].copy()
                for c in ['전월매출', '당월매출', '증감액']:
                    disp[c] = disp[c].apply(lambda x: f"{x:+,.0f}" if c == '증감액' else f"{x:,.0f}")
                disp['증감율(%)'] = disp['증감율(%)'].apply(lambda x: f"{x:+.1f}%")
                st.dataframe(disp, use_container_width=True)
            else:
                st.info("해당 조건의 부대가 없습니다.")

            # ══════════════════════════════════════════
            # 섹션 3: 급증 / 급감 부대 (30% 기준)
            # ══════════════════════════════════════════
            st.markdown("---")
            st.subheader(f"⚡ STEP 3 · 매출 급증 / 급감 부대 (전월 대비 ±{SURGE_THRESHOLD}% 이상)")
            st.caption(f"STEP 2의 증감 중에서 ±{SURGE_THRESHOLD}% 이상 변동한 부대만 추려낸 것입니다. 매출등급(STEP 1)과 함께 보면 어떤 등급의 부대가 급변했는지 바로 파악됩니다.")

            col_a3, col_b3 = st.columns([1, 2])
            with col_a3: filter_surge = st.selectbox("필터", ["전체", "📈 급증", "📉 급감", "➡️ 유지", "신규"], key="surge_filter")
            with col_b3: kw_surge = st.text_input("🔍 부대명 검색", key="kw_surge", placeholder="부대명 입력...")

            surge_df = unit_compare.copy()
            if filter_surge != "전체": surge_df = surge_df[surge_df['변동유형'] == filter_surge]
            surge_df = keyword_filter(surge_df, kw_surge).sort_values('증감율(%)', ascending=False)

            s1, s2, s3 = st.columns(3)
            s1.metric("📈 급증 부대", f"{len(unit_compare[unit_compare['변동유형']=='📈 급증'])}개")
            s2.metric("📉 급감 부대", f"{len(unit_compare[unit_compare['변동유형']=='📉 급감'])}개")
            s3.metric("📌 꾸준 저매출", f"{len(unit_compare[(unit_compare['변동유형']=='➡️ 유지') & (unit_compare['매출등급']=='🔵 저매출')])}개")

            if not surge_df.empty:
                surge_chart = surge_df[surge_df['변동유형'].isin(['📈 급증', '📉 급감'])].head(30)
                if not surge_chart.empty:
                    fig_surge = px.bar(surge_chart, x='부대명', y='증감율(%)',
                        color='매출등급',
                        color_discrete_map={'🔴 고매출': '#d9534f', '🟡 중매출': '#f0ad4e', '🔵 저매출': '#0275d8'},
                        title=f"급증/급감 부대 증감율 — 색상은 매출등급(STEP 1 기준)",
                        text=surge_chart['증감율(%)'].apply(lambda x: f"{x:+.1f}%"))
                    fig_surge.update_traces(textposition='outside')
                    fig_surge.update_layout(height=420)
                    fig_surge.update_xaxes(tickangle=-40)
                    st.plotly_chart(fig_surge, use_container_width=True)

                disp3 = surge_df[['부대명', '매출등급', '변동유형', '전월매출', '당월매출', '증감율(%)']].copy()
                disp3['전월매출']  = disp3['전월매출'].apply(lambda x: f"{x:,.0f}")
                disp3['당월매출']  = disp3['당월매출'].apply(lambda x: f"{x:,.0f}")
                disp3['증감율(%)'] = disp3['증감율(%)'].apply(lambda x: f"{x:+.1f}%")
                st.dataframe(disp3, use_container_width=True)
            else:
                st.info("해당 조건의 부대가 없습니다.")

            # ══════════════════════════════════════════
            # 섹션 4: 다음 달 영업 집중 부대 추천
            # ══════════════════════════════════════════
            st.markdown("---")
            st.subheader(f"🚀 STEP 4 · {selected_base_month + 1}월 영업 활동 집중 부대 추천")
            st.caption("STEP 1~3의 분석을 종합해 시스템이 자동 판단한 우선순위 추천입니다.")

            rec1 = unit_compare[
                (unit_compare['변동유형'] == '📉 급감') &
                (unit_compare['매출등급'].isin(['🔴 고매출', '🟡 중매출']))
            ].copy()
            rec1['추천사유'] = '🔥 고/중매출 급감 — 즉시 방문 및 원인 파악 필요'
            rec1['우선순위'] = 1

            rec2 = unit_compare[
                (unit_compare['변동유형'] == '📈 급증') &
                (unit_compare['매출등급'] == '🔵 저매출')
            ].copy()
            rec2['추천사유'] = '🌱 저매출 급증 — 성장 가능성 확인 및 관계 강화'
            rec2['우선순위'] = 2

            rec3 = unit_compare[
                (unit_compare['변동유형'] == '📈 급증') &
                (unit_compare['매출등급'].isin(['🔴 고매출', '🟡 중매출']))
            ].copy()
            rec3['추천사유'] = '💪 고/중매출 급증 — 모멘텀 유지 및 추가 제안'
            rec3['우선순위'] = 3

            rec4 = unit_compare[
                (unit_compare['변동유형'] == '➡️ 유지') &
                (unit_compare['매출등급'] == '🔵 저매출')
            ].copy()
            rec4['추천사유'] = '📌 꾸준한 저매출 유지 — 공략 가능성 점검 및 수요 발굴 필요'
            rec4['우선순위'] = 4

            rec_all = pd.concat([rec1, rec2, rec3, rec4]).drop_duplicates(subset='부대명')
            rec_all = rec_all.sort_values(['우선순위', '증감율(%)'])

            if not rec_all.empty:
                priority_colors = {1: '#d9534f', 2: '#5cb85c', 3: '#f0ad4e', 4: '#5bc0de'}
                priority_labels = {1: '🔥 즉시 대응', 2: '🌱 성장 관리', 3: '💪 모멘텀 유지', 4: '📌 꾸준 저매출 공략'}
                priority_desc = {
                    1: f"전월 대비 매출이 {SURGE_THRESHOLD}% 이상 급감했지만 연평균 기준 고/중매출 부대입니다. 기존 거래 관계가 있는 핵심 부대이므로 이탈 방지가 최우선입니다. 즉시 방문하여 감소 원인(경쟁사 진입, 담당자 교체, 불만 사항 등)을 파악하세요.",
                    2: f"연평균 기준 저매출이지만 이번 달 {SURGE_THRESHOLD}% 이상 급증했습니다. 아직 절대 매출은 낮지만 성장 신호가 포착된 부대입니다. 지금 집중 관리하면 중매출 이상으로 끌어올릴 수 있는 적기입니다.",
                    3: f"이미 고/중매출을 유지하면서 이번 달 추가로 {SURGE_THRESHOLD}% 이상 급증했습니다. 현재 영업 모멘텀이 가장 강한 부대입니다. 추가 품목 제안이나 납품 물량 확대를 적극 제안하세요.",
                    4: "급증/급감 없이 꾸준히 저매출을 유지하고 있는 부대입니다. 거래는 이어지고 있으나 잠재 수요가 충분히 발굴되지 않은 상태입니다. 품목 다양화 제안이나 담당자 관계 강화로 매출 확대를 노려볼 수 있습니다.",
                }

                for priority in [1, 2, 3, 4]:
                    grp = rec_all[rec_all['우선순위'] == priority]
                    if grp.empty: continue
                    color = priority_colors[priority]
                    label = priority_labels[priority]
                    desc  = priority_desc[priority]
                    st.markdown(f"""
                    <div style="border-left:5px solid {color};padding:12px 16px;margin-bottom:4px;
                                background-color:rgba(255,255,255,0.03);border-radius:4px;">
                        <strong style="color:{color};font-size:15px;">{label}</strong>
                        &nbsp;&nbsp;<span style="color:#aaa;font-size:13px;">({len(grp)}개 부대)</span>
                        <p style="margin:8px 0 0 0;font-size:13px;color:#ccc;line-height:1.6;">
                            💡 <b>추천 근거:</b> {desc}
                        </p>
                    </div>""", unsafe_allow_html=True)

                    disp_rec = grp[['부대명', '매출등급', '변동유형', '연매출', '당월매출', '증감율(%)', '추천사유']].copy()
                    disp_rec = disp_rec.rename(columns={'연매출': '연 누적 매출', '당월매출': f'{selected_base_month}월 매출'})
                    disp_rec['연 누적 매출']        = disp_rec['연 누적 매출'].apply(lambda x: f"{x:,.0f}")
                    disp_rec[f'{selected_base_month}월 매출'] = disp_rec[f'{selected_base_month}월 매출'].apply(lambda x: f"{x:,.0f}")
                    disp_rec['증감율(%)']            = disp_rec['증감율(%)'].apply(lambda x: f"{x:+.1f}%")
                    st.dataframe(disp_rec, use_container_width=True)
                    st.markdown("<br>", unsafe_allow_html=True)

                st.download_button(
                    label="📥 추천 부대 리스트 다운로드",
                    data=rec_all[['부대명', '매출등급', '변동유형', '연매출', '당월매출', '증감율(%)', '추천사유', '우선순위']]
                        .to_csv(index=False).encode('utf-8-sig'),
                    file_name=f'{selected_base_month+1}월_영업집중부대_추천.csv',
                    mime='text/csv'
                )
            else:
                st.info("추천 대상 부대가 없습니다. 데이터가 2개월 이상 있어야 분석이 가능합니다.")

    except Exception as e:
        import traceback
        st.error(f"🚨 오류 발생: {e}")
        st.code(traceback.format_exc())
