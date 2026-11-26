import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import numpy as np
from auto.units import UNIT_MASTER, MIN_DELIVERIES, normalize_unit, delivery_key, closed_through_month
from datetime import datetime, timedelta, timezone

def safe_sorted(iterable):
    return sorted([str(x) for x in iterable if str(x) not in ('nan', 'None', '')])

st.set_page_config(page_title="연간 매출 분석 대시보드", layout="wide", initial_sidebar_state="expanded")
st.title("📊 연간 매출 분석 대시보드")

st.sidebar.header("📁 데이터 연동 상태")

@st.cache_data(ttl=3600)
def load_excel_fallback():
    try:
        return pd.read_excel("data.xlsx", engine='openpyxl')
    except Exception as e:
        st.error(f"⚠️ 'data.xlsx' 파일을 읽을 수 없습니다: {e}")
        return None

def load_data():
    # 1순위: 구글 시트(자동화시스템) 자동 연동 / 실패 시 깃허브 data.xlsx로 대체
    try:
        has_secret = "SALES_API_URL" in st.secrets and "SALES_API_TOKEN" in st.secrets
    except Exception:
        has_secret = False
    if has_secret:
        try:
            from auto.sheet_loader import load_from_sheet
            d, closed = load_from_sheet()
            st.sidebar.success(f"✅ 자동화시스템 자동 연동 중 ({len(d):,}건)")
            if st.sidebar.button("🔄 지금 새로고침"):
                st.cache_data.clear()
                st.rerun()
            return d, closed
        except Exception as e:
            st.sidebar.warning(f"⚠️ 자동화시스템 연동 실패 → data.xlsx 사용\n\n{e}")
    else:
        st.sidebar.info("ℹ️ 자동화시스템 미설정 → data.xlsx 사용")
    st.sidebar.success("✅ 깃허브 data.xlsx 연동 중")
    return load_excel_fallback(), None

raw_df, sheet_closed = load_data()

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
        df['부대명'] = df['부대명'].map(normalize_unit)   # 하위 부대(51군지단/31사단 등) -> 상위 부대로 통합

        df = df[pd.to_numeric(df['매출'].astype(str).str.replace(',', '', regex=False), errors='coerce').notna()]
        df['구분'] = df['구분'].replace({'nan': '미분류', 'None': '미분류', '': '미분류'})   # 구분 공란도 합계에 포함
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

        # ── 부대별 분석 대상: 부대별정보에 있는 부대 중 연간 납품 3회 이상 ──
        # (합계 매출에는 모든 거래처 포함 / 부대별 세분화에서만 제외)
        df['_납품키'] = df['납품일'].map(delivery_key)
        in_master = df[df['부대명'].isin(UNIT_MASTER)]
        delivery_cnt = in_master.groupby('부대명')['_납품키'].nunique()
        valid_unit_names = delivery_cnt[delivery_cnt >= MIN_DELIVERIES].index
        df_unit = in_master[in_master['부대명'].isin(valid_unit_names)]

        not_master = df[~df['부대명'].isin(UNIT_MASTER)].groupby('부대명')['매출'].sum()
        few_delivery = delivery_cnt[delivery_cnt < MIN_DELIVERIES]
        _closed = sheet_closed if sheet_closed is not None else closed_through_month()
        st.sidebar.markdown("---")
        st.sidebar.caption(f"🗓️ {_closed}월까지(마감): 확정 매출 · {_closed + 1}월부터(미마감): 확정+예정")
        with st.sidebar.expander(f"📌 부대별 분석 대상 {len(valid_unit_names)}개 부대"):
            st.caption(
            f"- 합계 매출에만 포함(부대별정보에 없음): "
            + (", ".join(f"{k}({v:,.0f}원)" for k, v in not_master.items()) or "없음") + "\n"
            f"- 부대별 분석 제외(연간 납품 {MIN_DELIVERIES}회 미만): "
            + (", ".join(f"{k}({v}회)" for k, v in few_delivery.items()) or "없음")
            )

        KST = timezone(timedelta(hours=9))
        NOW_M = datetime.now(KST).month
        TARGET_M = min(NOW_M + 1, 12)   # 기준(당월) = 다음 달 예정월 (예: 11월 현재 → 12월)
        months_all = sorted(int(m) for m in df['월'].unique())
        HI, MID, LO = '🔴 상위', '⚪ 중위', '🔵 하위'
        GRADE_BG = {HI: 'background-color: rgba(229,83,75,0.22)', LO: 'background-color: rgba(59,142,234,0.22)'}


        def man(x):
            """금액을 '1.2억' / '3,456만' 형태로 짧게 표기."""
            return f"{x/1e8:,.1f}억" if abs(x) >= 1e8 else f"{x/1e4:,.0f}만"


        def styled(d, num_cols=(), pct_cols=()):
            fmt = {c: "{:,.0f}" for c in num_cols}
            fmt.update({c: "{:+.1f}%" for c in pct_cols})
            sty = d.style.format(fmt, na_rep='-')
            if '등급' in d.columns:
                sty = sty.apply(lambda r: [GRADE_BG.get(r['등급'], '')] * len(r), axis=1)
            return sty


        def show_table(d, num_cols=(), pct_cols=(), height=None, spark=None):
            """숫자형을 유지한 채 표시 → 제목 클릭 시 오름/내림차순 정렬이 숫자 기준으로 동작."""
            kw = dict(width='stretch', hide_index=True)
            if height:
                kw['height'] = height
            if spark:
                kw['column_config'] = {spark: st.column_config.LineChartColumn(spark, width='small')}
            st.dataframe(styled(d, num_cols, pct_cols), **kw)


        unit_pivot_raw = df_unit.pivot_table(index='부대명', columns='월', values='매출', aggfunc='sum', fill_value=0)
        unit_pivot_raw = unit_pivot_raw.reindex(columns=months_all, fill_value=0)
        unit_pivot_raw.columns = [int(c) for c in unit_pivot_raw.columns]

        tab1, tab2 = st.tabs(["📊 매출 현황", "🎯 영업 전략 인사이트"])

        # ══════════════════════════════════════════
        # TAB 1 : 매출 현황
        # ══════════════════════════════════════════
        with tab1:
            total = df['매출'].sum()
            monthly = df.groupby('월')['매출'].sum().reindex(months_all, fill_value=0)
            cur_m = max([m for m in months_all if m <= TARGET_M] or [months_all[-1]])
            k1, k2, k3 = st.columns(3)
            k1.metric("💰 올해 누적 매출", man(total), help=f"{total:,.0f}원")
            k2.metric(f"📅 {cur_m}월 매출" + (" (예정 포함)" if cur_m > _closed else ""), man(monthly.get(cur_m, 0)), help=f"{monthly.get(cur_m, 0):,.0f}원")
            k3.metric("📈 월평균 매출", man(monthly[monthly.index <= cur_m].mean()))

            mdf = monthly.reset_index()
            mdf.columns = ['월', '매출']
            mdf['구분'] = np.where(mdf['월'] <= _closed, '확정(마감)', '확정+예정(미마감)')
            fig_m = px.bar(mdf, x='월', y='매출', color='구분', text=mdf['매출'].map(man),
                           color_discrete_map={'확정(마감)': '#4C78A8', '확정+예정(미마감)': '#9ecae9'})
            fig_m.update_traces(textposition='outside', cliponaxis=False)
            fig_m.update_layout(height=300, margin=dict(l=10, r=10, t=30, b=10), yaxis_title=None,
                                legend=dict(orientation='h', y=1.12, x=0, title=None))
            fig_m.update_xaxes(dtick=1, ticksuffix='월', title=None)
            fig_m.update_yaxes(showticklabels=False)
            st.plotly_chart(fig_m, width='stretch')

            # ── 품목 순위 (상위 10 / 하위 10 한 줄) ──
            st.markdown("---")
            st.subheader("🏆 품목 매출 순위")
            st.caption("막대를 클릭하면 아래에 해당 품목의 월별 매출이 표시됩니다.")
            items = df.groupby(['구분', '품목'])['매출'].sum().reset_index()
            items = items[items['매출'] > 0]
            items['라벨'] = items['품목'] + ' (' + items['구분'] + ')'
            item_month = df.groupby(['구분', '품목', '월'])['매출'].sum()

            def item_chart(data, color, key, title):
                st.markdown(f"**{title}**")
                fig = px.bar(data, x='매출', y='라벨', orientation='h', text=data['매출'].map(man),
                             custom_data=['구분', '품목'], color_discrete_sequence=[color])
                fig.update_traces(textposition='outside', cliponaxis=False)
                fig.update_layout(height=360, margin=dict(l=5, r=50, t=5, b=5), xaxis_title=None, yaxis_title=None,
                                  xaxis=dict(showticklabels=False, range=[0, data['매출'].max() * 1.25]))
                ev = st.plotly_chart(fig, width='stretch', on_select='rerun', selection_mode='points', key=key)
                pts = (ev or {}).get('selection', {}).get('points', [])
                if pts and pts[0].get('customdata'):
                    cat, name = pts[0]['customdata'][:2]
                    d = item_month.loc[(cat, name)].reindex(months_all, fill_value=0).reset_index()
                    d.columns = ['월', '매출']
                    f2 = px.bar(d, x='월', y='매출', text=d['매출'].map(lambda v: man(v) if v else ''),
                                color_discrete_sequence=[color], title=f"{name} 월별 매출")
                    f2.update_traces(textposition='outside', cliponaxis=False)
                    f2.update_layout(height=230, margin=dict(l=5, r=5, t=35, b=5), yaxis=dict(showticklabels=False, title=None))
                    f2.update_xaxes(dtick=1, ticksuffix='월', title=None)
                    st.plotly_chart(f2, width='stretch', key=key + '_d')

            cL, cR = st.columns(2)
            with cL:
                item_chart(items.nlargest(10, '매출').sort_values('매출'), '#e5534b', 'top10', '🔴 상위 10개 품목')
            with cR:
                item_chart(items.nsmallest(10, '매출').sort_values('매출', ascending=False), '#3b8eea', 'bottom10', '🔵 하위 10개 품목')

            # ── 구분별 카드 ──
            st.markdown("---")
            st.subheader("🗂️ 구분별 매출")
            cat_sum = df.groupby('구분')['매출'].sum().sort_values(ascending=False)
            cards = "".join(
                f"<div style='border:1px solid rgba(128,128,128,.35);border-radius:8px;padding:10px 12px;'>"
                f"<div style='font-size:13px;opacity:.75'>{c}</div>"
                f"<div style='font-size:20px;font-weight:700'>{man(v)}</div>"
                f"<div style='font-size:12px;opacity:.6'>{v / total * 100:.1f}%</div></div>"
                for c, v in cat_sum.items())
            st.markdown(f"<div style='display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:8px;'>{cards}</div>",
                        unsafe_allow_html=True)
            sel_cat = st.pills("상세 보기할 구분 선택", options=list(cat_sum.index), key='cat_pill')
            if sel_cat:
                cdf = df[df['구분'] == sel_cat]
                pv = cdf.pivot_table(index='품목', columns='월', values='매출', aggfunc='sum', fill_value=0)
                pv.columns = [f"{int(c)}월" for c in pv.columns]
                pv.insert(0, '합계', pv.sum(axis=1))
                pv = pv.sort_values('합계', ascending=False).reset_index()
                show_table(pv, num_cols=[c for c in pv.columns if c != '품목'])
                d1, d2 = st.columns(2)
                s_item = d1.selectbox("품목", pv['품목'], key='drill_item')
                s_mon = d2.selectbox("월", sorted(cdf['월'].unique()), format_func=lambda x: f"{x}월", key='drill_mon')
                dd = cdf[(cdf['품목'] == s_item) & (cdf['월'] == s_mon)][['부대명', '납품일', '수량', '매출']]
                if dd.empty:
                    st.caption(f"{s_mon}월 '{s_item}' 발주 내역 없음")
                else:
                    show_table(dd, num_cols=['수량', '매출'])

            # ── 상세 검색 ──
            st.markdown("---")
            with st.expander("🔍 전체 데이터 상세 검색"):
                f1, f2, f3, f4 = st.columns(4)
                fu = f1.multiselect("부대명", safe_sorted(df['부대명'].unique()))
                fc = f2.multiselect("구분", safe_sorted(df['구분'].unique()))
                fi = f3.multiselect("품목", safe_sorted(df['품목'].unique()))
                fm = f4.multiselect("월", months_all, format_func=lambda x: f"{x}월")
                q = df
                if fu: q = q[q['부대명'].isin(fu)]
                if fc: q = q[q['구분'].isin(fc)]
                if fi: q = q[q['품목'].isin(fi)]
                if fm: q = q[q['월'].isin(fm)]
                q = q[['납품일', '부대명', '구분', '품목', '수량', '단가(Vat별도)', '매출']]
                st.caption(f"{len(q):,}건 · 합계 {q['매출'].sum():,.0f}원")
                show_table(q, num_cols=['수량', '단가(Vat별도)', '매출'], height=400)
                st.download_button("📥 검색 결과 다운로드", q.to_csv(index=False).encode('utf-8-sig'),
                                   file_name='매출_상세검색.csv', mime='text/csv')

        # ══════════════════════════════════════════
        # TAB 2 : 영업 전략 인사이트
        # ══════════════════════════════════════════
        with tab2:
            months_desc = sorted(months_all, reverse=True)
            t1, t0, t2 = st.columns([1, 1, 3])
            base = t1.selectbox("📅 기준 월", months_desc,
                                index=next((i for i, m in enumerate(months_desc) if m <= TARGET_M), 0),
                                format_func=lambda x: f"{x}월")
            prev = base - 1
            n_show = t0.selectbox("🗓️ 표시 기간", [3, 6, 12], index=0, format_func=lambda x: f"최근 {x}개월")
            t2.markdown(
                f"<div style='padding-top:30px;font-size:13px;opacity:.8'>행 색상: "
                f"<span style='background:rgba(229,83,75,.3);padding:2px 6px;border-radius:4px'>🔴 상위</span> 누적 매출 상위 1/3 · "
                f"<span style='background:rgba(59,142,234,.3);padding:2px 6px;border-radius:4px'>🔵 하위</span> 하위 1/3 "
                f"(1월~{base}월 누적 기준)</div>", unsafe_allow_html=True)

            upto = [m for m in months_all if m <= base]
            up = unit_pivot_raw[upto]
            cum = up.sum(axis=1)
            WSHOW = upto[-n_show:]                     # 표에 보여줄 월 (기본 최근 3개월)
            wcols = [f"{m}월" for m in WSHOW]
            if len(cum) >= 3:
                grade = pd.qcut(cum.rank(method='first'), q=3, labels=[LO, MID, HI]).astype(str)
            else:
                grade = pd.Series(MID, index=cum.index)
            cur = unit_pivot_raw.get(base, pd.Series(0, index=unit_pivot_raw.index))
            prv = unit_pivot_raw.get(prev, pd.Series(0, index=unit_pivot_raw.index))

            # ① 부대별 매출 현황
            st.subheader(f"① 부대별 매출 현황 ({base}월 기준)")
            t = pd.DataFrame({'부대명': cum.index, '등급': grade.values, '누적매출': cum.values,
                              '월평균': (cum / max(len(upto), 1)).values})
            for m in WSHOW:
                t[f"{m}월"] = up[m].values
            t['추이'] = [list(up.loc[u].values) for u in cum.index]
            t = t[t['누적매출'] > 0].sort_values('누적매출', ascending=False)
            show_table(t, num_cols=['누적매출', '월평균'] + wcols,
                       height=420, spark='추이')

            # ② 이번 달 급증 / 급감 (당월 vs 직전 3개월 평균)
            st.markdown("---")
            SURGE = 30
            prev3 = [m for m in upto if m < base][-3:]
            st.subheader(f"② {base}월 급증 / 급감 부대")
            st.caption(f"{base}월 매출을 직전 3개월({', '.join(f'{m}월' for m in prev3)}) 평균과 비교해 ±{SURGE}% 이상 변한 부대"
                       + (" · 미마감 월은 예정 포함" if base > _closed else ""))
            if not prev3:
                st.info("비교할 직전 월 데이터가 없습니다.")
            else:
                sv = pd.DataFrame({'부대명': cum.index, '등급': grade.values})
                for m in prev3:
                    sv[f'{m}월'] = up[m].values
                sv['3개월 평균'] = up[prev3].mean(axis=1).values
                sv[f'{base}월'] = cur.reindex(cum.index).fillna(0).values
                sv['증감액'] = sv[f'{base}월'] - sv['3개월 평균']
                sv['증감률'] = np.where(sv['3개월 평균'] > 0, sv['증감액'] / sv['3개월 평균'].replace(0, np.nan) * 100, np.nan)
                sv['비고'] = ''
                sv.loc[(sv['3개월 평균'] == 0) & (sv[f'{base}월'] > 0), '비고'] = '신규'
                sv.loc[(sv['3개월 평균'] > 0) & (sv[f'{base}월'] == 0), '비고'] = '발주 없음'
                up_df = sv[(sv['증감률'] >= SURGE) | (sv['비고'] == '신규')].sort_values('증감액', ascending=False)
                dn_df = sv[sv['증감률'] <= -SURGE].sort_values('증감액')
                scols = ['부대명', '등급'] + [f'{m}월' for m in prev3] + ['3개월 평균', f'{base}월', '증감률', '비고']
                snum = [f'{m}월' for m in prev3] + ['3개월 평균', f'{base}월']
                s1, s2 = st.columns(2)
                with s1:
                    st.markdown(f"**📈 급증 ({len(up_df)}개)**")
                    show_table(up_df[scols], num_cols=snum, pct_cols=['증감률'])
                with s2:
                    st.markdown(f"**📉 급감 ({len(dn_df)}개)**")
                    show_table(dn_df[scols], num_cols=snum, pct_cols=['증감률'])

            # ③ 분기 분석
            st.markdown("---")
            st.subheader("③ 분기 분석")
            qmap = {1: [1, 2, 3], 2: [4, 5, 6], 3: [7, 8, 9], 4: [10, 11, 12]}
            qs = [q for q, ms in qmap.items() if any(m in months_all for m in ms)]
            qcol1, qcol2 = st.columns([1, 4])
            qsel = qcol1.selectbox("분기", qs, index=len(qs) - 1, format_func=lambda q: f"{q}분기")
            qms = [m for m in qmap[qsel] if m in unit_pivot_raw.columns]
            pms = [m for m in qmap.get(qsel - 1, []) if m in unit_pivot_raw.columns]
            plan_note = " · 미마감 월은 예정 포함" if any(m > _closed for m in qms) else ""
            qcol2.markdown(f"<div style='padding-top:32px;font-size:13px;opacity:.8'>"
                           f"{qsel}분기({qms[0]}~{qms[-1]}월)를 {'전 분기와 비교' if pms else '단독 분석'}{plan_note}</div>",
                           unsafe_allow_html=True)
            QTH, SWING = 10, 50
            qd = pd.DataFrame({'부대명': unit_pivot_raw.index})
            qd['등급'] = qd['부대명'].map(grade).fillna(MID)
            for m in qms:
                qd[f'{m}월'] = unit_pivot_raw[m].values
            qd[f'{qsel}분기'] = unit_pivot_raw[qms].sum(axis=1).values
            qnum = [f'{m}월' for m in qms] + [f'{qsel}분기']
            if pms:
                qd[f'{qsel - 1}분기'] = unit_pivot_raw[pms].sum(axis=1).values
                qd['증감률'] = np.where(qd[f'{qsel - 1}분기'] > 0,
                                     (qd[f'{qsel}분기'] - qd[f'{qsel - 1}분기']) / qd[f'{qsel - 1}분기'].replace(0, np.nan) * 100, np.nan)
                qcols = ['부대명', '등급', f'{qsel - 1}분기'] + qnum + ['증감률']
                rise = qd[qd['증감률'] >= QTH].sort_values('증감률', ascending=False)
                fall = qd[qd['증감률'] <= -QTH].sort_values('증감률')
                r1, r2 = st.columns(2)
                with r1:
                    st.markdown(f"**📈 상승 부대 ({len(rise)}개)** · 전 분기 대비 +{QTH}% 이상")
                    show_table(rise[qcols], num_cols=[f'{qsel - 1}분기'] + qnum, pct_cols=['증감률'])
                with r2:
                    st.markdown(f"**📉 하락 부대 ({len(fall)}개)** · 전 분기 대비 −{QTH}% 이상")
                    show_table(fall[qcols], num_cols=[f'{qsel - 1}분기'] + qnum, pct_cols=['증감률'])
            qavg = qd[[f'{m}월' for m in qms]].mean(axis=1)
            qd['등락폭'] = np.where(qavg > 0, (qd[[f'{m}월' for m in qms]].max(axis=1) - qd[[f'{m}월' for m in qms]].min(axis=1)) / qavg.replace(0, np.nan) * 100, np.nan)
            swing = qd[(qd['등락폭'] >= SWING) & (qd[f'{qsel}분기'] > 0)].sort_values('등락폭', ascending=False)
            st.markdown(f"**🌊 등락폭이 큰 부대 ({len(swing)}개)** · 분기 내 (최고월 − 최저월) ÷ 월평균 ≥ {SWING}% → 발주가 불규칙한 부대")
            show_table(swing[['부대명', '등급'] + qnum + ['등락폭']], num_cols=qnum, pct_cols=['등락폭'])

    except Exception as e:
        import traceback
        st.error(f"🚨 오류 발생: {e}")
        st.code(traceback.format_exc())
