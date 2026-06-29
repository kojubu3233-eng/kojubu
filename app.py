import streamlit as st
import pandas as pd
import plotly.express as px
import re

# 정렬 시 타입 혼재(float+str) 오류 방지용 헬퍼 함수
def safe_sorted(iterable):
    return sorted([str(x) for x in iterable if str(x) not in ('nan', 'None', '')])

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(page_title="영업 매출 분석 대시보드", layout="wide", initial_sidebar_state="expanded")

st.title("📊 자사 영업 매출 분석 대시보드")
st.markdown("매월 엑셀 데이터를 업로드하여 **중점 영업 전략**을 도출하세요.")
st.markdown("---")

# 사이드바 설정
st.sidebar.header("📁 데이터 초기화/업로드")
uploaded_file = st.sidebar.file_uploader("월 매출 파일을 업로드해 주세요 (엑셀 또는 CSV).", type=['xlsx', 'xls', 'csv'])

if uploaded_file:
    try:
        # 2. 파일 형식(CSV/Excel)에 맞게 안전하게 로드
        if uploaded_file.name.endswith('.csv'):
            try:
                df = pd.read_csv(uploaded_file, encoding='cp949')
            except UnicodeDecodeError:
                df = pd.read_csv(uploaded_file, encoding='utf-8')
        else:
            df = pd.read_excel(uploaded_file, header=0)

        # 3. 코드용 표준 이름으로 강제 재설정
        if len(df.columns) >= 7:
            df.columns = list(df.columns[:7])
            df.columns = ['납품일', '부대명', '구분', '품목', '수량', '단가(Vat별도)', '매출']
        else:
            st.error(f"⚠️ 업로드된 파일의 열(Column) 개수가 부족합니다. (현재 {len(df.columns)}개, 최소 7개 필요)")
            st.stop()

        # 4. 모든 텍스트 컬럼을 str로 통일
        for col in ['납품일', '부대명', '구분', '품목']:
            df[col] = df[col].astype(str).str.strip()

        # ✅ 헤더가 데이터로 混入된 행 제거 (매출 컬럼이 숫자가 아닌 행 삭제)
        df = df[pd.to_numeric(df['매출'].astype(str).str.replace(',', '', regex=False), errors='coerce').notna()]
        # ✅ nan/None 텍스트 행 제거
        df = df[~df['구분'].isin(['nan', 'None', ''])]
        df = df[~df['부대명'].isin(['nan', 'None', ''])]

        for col in ['수량', '단가(Vat별도)', '매출']:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 5. 월 추출
        def extract_month_safely(val):
            if pd.isna(val):
                return 1
            val_str = str(val).strip()
            match = re.search(r'(\d+)월', val_str)
            if match:
                return int(match.group(1))
            match2 = re.search(r'-(\d{1,2})-', val_str)
            if match2:
                return int(match2.group(1))
            # 엑셀 날짜 직렬번호 처리
            try:
                ts = pd.to_datetime(val_str, errors='coerce')
                if pd.notna(ts):
                    return ts.month
            except:
                pass
            return 1

        df['월'] = df['납품일'].apply(extract_month_safely).astype(int)
        df = df[df['월'].between(1, 12)]  # 유효한 월만 유지
        df = df.sort_values('월')

        # ============================================================
        # 📊 [화면 1] 전체 매출 현황
        # ============================================================
        st.header("1. 올해 전체 매출 현황 및 요약")

        total_sales = df['매출'].sum()
        st.metric(label="💰 올해 누적 총 매출액", value=f"{total_sales:,.0f} 원")

        monthly_total = df.groupby('월')['매출'].sum().reset_index()
        fig_total = px.line(
            monthly_total,
            x='월',
            y='매출',
            markers=True,
            title="📅 월별 전체 매출 추이",
            text=monthly_total['매출'].apply(lambda x: f"{x:,.0f}원" if x > 0 else "")
        )
        fig_total.update_traces(textposition="top center", line=dict(width=3, color="#1f77b4"))
        fig_total.update_xaxes(dtick=1, title="월")
        fig_total.update_yaxes(title="매출액 (원)")
        st.plotly_chart(fig_total, use_container_width=True)

        # ============================================================
        # 📊 계층형 데이터 요약: (대분류) 구분별 총매출 → (소분류) 품목별 피벗
        # ============================================================
        st.subheader("💡 계층형 데이터 요약 (구분별 ➔ 품목별)")
        st.caption("**[구분]** 항목을 클릭하면 해당 구분의 **구분 총매출**과 **품목별 월별 매출 피벗**을 확인할 수 있습니다.")

        # 구분별 총매출 미리 계산 (매출 내림차순 정렬)
        category_summary = df.groupby('구분')['매출'].sum().reset_index()
        category_summary.columns = ['구분', '총매출']
        category_summary = category_summary.sort_values('총매출', ascending=False)

        for _, cat_row in category_summary.iterrows():
            category = cat_row['구분']
            cat_total = cat_row['총매출']

            # ── 대분류 expander: 구분명 + 총매출 표시 ──
            with st.expander(f"📁 {category}  |  총 매출: {cat_total:,.0f}원  (클릭하여 품목별 상세 보기)"):

                cat_df = df[df['구분'] == category]

                # ── 소분류: 품목별 총매출 요약 바 차트 ──
                item_summary = cat_df.groupby('품목')['매출'].sum().reset_index()
                item_summary.columns = ['품목', '총매출']
                item_summary = item_summary.sort_values('총매출', ascending=False)

                fig_bar = px.bar(
                    item_summary,
                    x='품목',
                    y='총매출',
                    text=item_summary['총매출'].apply(lambda x: f"{x:,.0f}원"),
                    title=f"[{category}] 품목별 총 매출",
                    color='총매출',
                    color_continuous_scale='Blues'
                )
                fig_bar.update_traces(textposition='outside')
                fig_bar.update_layout(showlegend=False, coloraxis_showscale=False, height=350)
                st.plotly_chart(fig_bar, use_container_width=True)

                # ── 소분류 피벗: 품목 × 월별 매출 ──
                st.markdown("**📋 품목별 월별 매출 상세 (피벗)**")
                pivot_df = pd.pivot_table(
                    cat_df,
                    values='매출',
                    index=['품목'],
                    columns=['월'],
                    aggfunc='sum',
                    fill_value=0
                )
                pivot_df.columns = [f"{int(c)}월" for c in pivot_df.columns]
                pivot_df['총합계'] = pivot_df.sum(axis=1)
                pivot_df = pivot_df.sort_values('총합계', ascending=False)

                safe_pivot_df = pivot_df.copy()
                for col in safe_pivot_df.columns:
                    safe_pivot_df[col] = safe_pivot_df[col].apply(lambda x: f"{x:,.0f}")

                st.dataframe(safe_pivot_df, use_container_width=True)

        st.markdown("---")

        # ============================================================
        # 📈 [화면 2] 상세 분석 · 부대별 현황 · 증감율 (MoM)
        # ============================================================
        st.header("2. 매출 상세 분석 및 부대별 현황")

        # ----------------------------------------------------
        # 2-1. 부대별 전체 매출 현황 (마스터 리스트)
        # ----------------------------------------------------
        st.subheader("🏢 부대별 전체 매출 현황 (마스터 리스트)")

        # 부대명 기준, 월별 매출 피벗 생성
        unit_pivot = pd.pivot_table(
            df,
            index='부대명',
            columns='월',
            values='매출',
            aggfunc='sum',
            fill_value=0
        )

        # 컬럼명 'N월' 형식으로 변환
        unit_pivot.columns = [f"{int(c)}월" for c in unit_pivot.columns]

        # '총매출' 열 추가 및 부대명 가나다 오름차순 정렬
        unit_pivot['총매출'] = unit_pivot.sum(axis=1)
        unit_pivot = unit_pivot.sort_index(ascending=True)

        # 포맷팅 (콤마 적용)
        safe_unit_pivot = unit_pivot.map(lambda x: f"{x:,.0f}")

        st.dataframe(safe_unit_pivot, use_container_width=True)

        st.download_button(
            label="📥 부대별 현황 엑셀 다운로드",
            data=unit_pivot.to_csv().encode('utf-8-sig'),
            file_name='부대별_매출_현황.csv',
            mime='text/csv'
        )

        st.markdown("---")

        # ----------------------------------------------------
        # 2-2. 상세 필터 및 전월 대비 증감율 분석 (MoM)
        # ----------------------------------------------------
        st.subheader("📈 상세 필터 및 전월 대비 증감율 분석")

        col1, col2, col3 = st.columns(3)
        with col1:
            selected_units = st.multiselect("📍 부대명 선택", options=safe_sorted(df['부대명'].unique()))
        with col2:
            selected_categories = st.multiselect("🏷️ 구분 선택", options=safe_sorted(df['구분'].unique()))
        with col3:
            selected_months = st.multiselect("📅 월 선택", options=sorted(df['월'].unique()))

        analysis_df = df.copy()
        if selected_units:
            analysis_df = analysis_df[analysis_df['부대명'].isin(selected_units)]
        if selected_categories:
            analysis_df = analysis_df[analysis_df['구분'].isin(selected_categories)]
        if selected_months:
            analysis_df = analysis_df[analysis_df['월'].isin(selected_months)]

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
                / filtered_monthly.loc[has_prev, '전월매출']
            ) * 100

            # 차트 제목: 선택한 필터 조건을 동적으로 반영
            title_parts = []
            if selected_units:
                # 부대명이 많으면 앞 3개만 표시하고 나머지는 '외 N개'
                if len(selected_units) <= 3:
                    title_parts.append(" · ".join(selected_units))
                else:
                    title_parts.append(f"{' · '.join(selected_units[:3])} 외 {len(selected_units)-3}개")
            if selected_categories:
                title_parts.append(" · ".join(selected_categories))
            if selected_months:
                title_parts.append(f"{'/'.join(str(m)+'월' for m in sorted(selected_months))}")

            if title_parts:
                chart_title = "📈 [ " + "  |  ".join(title_parts) + " ] 매출액 추이"
            else:
                chart_title = "📈 전체 매출액 추이"

            fig_f = px.line(
                filtered_monthly,
                x='월',
                y='매출',
                markers=True,
                title=chart_title,
                text=filtered_monthly['매출'].apply(lambda x: f"{x:,.0f}원" if x > 0 else "")
            )
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
                        s_html, b_color = "<span style='color:gray; font-weight:bold;'>- 0%</span>", "gray"
                    elif prev_row['매출'] == 0 and row['매출'] > 0:
                        s_html, b_color = "<span style='color:#d9534f; font-weight:bold;'>▲ 신규 발생</span>", "#d9534f"
                    else:
                        if pct > 0:
                            s_html, b_color = f"<span style='color:#d9534f; font-weight:bold;'>▲ {pct:.1f}% 상승</span>", "#d9534f"
                        elif pct < 0:
                            s_html, b_color = f"<span style='color:#0275d8; font-weight:bold;'>▼ {abs(pct):.1f}% 하락</span>", "#0275d8"
                        else:
                            s_html, b_color = "<span style='color:gray; font-weight:bold;'>- 변동 없음</span>", "gray"

                    with m_cols[m_idx]:
                        st.markdown(
                            f"""
                            <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px;
                                        border-left: 5px solid {b_color}; text-align: center; margin-bottom: 12px;">
                                <p style="margin: 0; font-size: 13px; color: #6c757d;">{p_mon}월 ➔ {c_mon}월</p>
                                <p style="margin: 6px 0 0 0; font-size: 16px;">{s_html}</p>
                            </div>
                            """, unsafe_allow_html=True)

            # ============================================================
            # 🔍 [화면 3] 엑셀식 상세 검색
            # ============================================================
            st.markdown("---")
            st.subheader("📋 전체/선택 데이터 엑셀식 상세 검색")
            st.caption("아래의 다중 필터(엑셀의 자동 필터 역할)를 이용해 표의 데이터를 자유롭게 걸러내세요.")

            with st.expander("🔍 상세 검색 및 데이터 표 펼치기", expanded=False):
                f_col1, f_col2, f_col3, f_col4 = st.columns(4)
                with f_col1:
                    raw_date = st.multiselect("☑️ 납품일", options=safe_sorted(analysis_df['납품일'].unique()))
                with f_col2:
                    raw_unit = st.multiselect("☑️ 부대명", options=safe_sorted(analysis_df['부대명'].unique()))
                with f_col3:
                    raw_cat = st.multiselect("☑️ 구분", options=safe_sorted(analysis_df['구분'].unique()))
                with f_col4:
                    raw_item = st.multiselect("☑️ 품목", options=safe_sorted(analysis_df['품목'].unique()))

                final_table_df = analysis_df.copy()
                if raw_date:
                    final_table_df = final_table_df[final_table_df['납품일'].isin(raw_date)]
                if raw_unit:
                    final_table_df = final_table_df[final_table_df['부대명'].isin(raw_unit)]
                if raw_cat:
                    final_table_df = final_table_df[final_table_df['구분'].isin(raw_cat)]
                if raw_item:
                    final_table_df = final_table_df[final_table_df['품목'].isin(raw_item)]

                st.write(f"결과: **{len(final_table_df)}건**")

                display_df = final_table_df.copy()
                for col in ['수량', '단가(Vat별도)', '매출']:
                    display_df[col] = display_df[col].apply(lambda x: f"{x:,.0f}")

                st.dataframe(display_df, use_container_width=True)

        else:
            st.warning("선택하신 필터 조건과 일치하는 데이터가 없습니다.")

    except Exception as e:
        import traceback
        st.error(f"🚨 파일을 처리하는 중 예기치 못한 오류가 발생했습니다: {e}")
        st.code(traceback.format_exc())  # 개발 중 디버깅용 전체 스택 출력