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
            d = load_from_sheet()
            st.sidebar.success(f"✅ 자동화시스템 자동 연동 중 ({len(d):,}건)")
            if st.sidebar.button("🔄 지금 새로고침"):
                st.cache_data.clear()
                st.rerun()
            return d
        except Exception as e:
            st.sidebar.warning(f"⚠️ 자동화시스템 연동 실패 → data.xlsx 사용\n\n{e}")
    else:
        st.sidebar.info("ℹ️ 자동화시스템 미설정 → data.xlsx 사용")
    st.sidebar.success("✅ 깃허브 data.xlsx 연동 중")
    return load_excel_fallback()

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
        try:
            _closed = closed_through_month(st.secrets.get("CLOSED_THROUGH_MONTH"))
        except Exception:
            _closed = closed_through_month()
        st.sidebar.markdown("---")
        st.sidebar.caption(f"🗓️ 마감 ~{_closed}월: 확정 / {_closed + 1}월~: 확정+예정")
        with st.sidebar.expander(f"📌 부대별 분석 대상 {len(valid_unit_names)}개 부대"):
            st.caption(
            f"- 합계 매출에만 포함(부대별정보에 없음): "
            + (", ".join(f"{k}({v:,.0f}원)" for k, v in not_master.items()) or "없음") + "\n"
            f"- 부대별 분석 제외(연간 납품 {MIN_DELIVERIES}회 미만): "
            + (", ".join(f"{k}({v}회)" for k, v in few_delivery.items()) or "없음")
            )

        KST = timezone(timedelta(hours=9))
        NOW_M = datetime.now(KST).month
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
            cur_m = NOW_M if NOW_M in monthly.index else months_all[-1]
            k1, k2, k3 = st.columns(3)
            k1.metric("💰 올해 누적 매출", man(total), help=f"{total:,.0f}원")
            k2.metric(f"📅 {cur_m}월 매출", man(monthly.get(cur_m, 0)), help=f"{monthly.get(cur_m, 0):,.0f}원")
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
            t1, t2 = st.columns([1, 3])
            base = t1.selectbox("📅 기준 월", months_desc,
                                index=next((i for i, m in enumerate(months_desc) if m <= NOW_M), 0),
                                format_func=lambda x: f"{x}월")
            prev = base - 1
            t2.markdown(
                f"<div style='padding-top:30px;font-size:13px;opacity:.8'>행 색상: "
                f"<span style='background:rgba(229,83,75,.3);padding:2px 6px;border-radius:4px'>🔴 상위</span> 누적 매출 상위 1/3 · "
                f"<span style='background:rgba(59,142,234,.3);padding:2px 6px;border-radius:4px'>🔵 하위</span> 하위 1/3 "
                f"(1월~{base}월 누적 기준)</div>", unsafe_allow_html=True)

            upto = [m for m in months_all if m <= base]
            up = unit_pivot_raw[upto]
            cum = up.sum(axis=1)
            if len(cum) >= 3:
                grade = pd.qcut(cum.rank(method='first'), q=3, labels=[LO, MID, HI]).astype(str)
            else:
                grade = pd.Series(MID, index=cum.index)
            cur = unit_pivot_raw.get(base, pd.Series(0, index=unit_pivot_raw.index))
            prv = unit_pivot_raw.get(prev, pd.Series(0, index=unit_pivot_raw.index))

            # ① 부대별 매출 현황
            st.subheader(f"① 부대별 매출 현황 ({base}월 기준)")
            last3 = [m for m in upto[-3:]]
            t = pd.DataFrame({'부대명': cum.index, '등급': grade.values, '누적매출': cum.values,
                              '월평균': (cum / max(len(upto), 1)).values})
            for m in last3:
                t[f"{m}월" + (" (당월)" if m == base else "")] = up[m].values
            t['추이'] = [list(up.loc[u].values) for u in cum.index]
            t = t.sort_values('누적매출', ascending=False)
            show_table(t, num_cols=['누적매출', '월평균'] + [c for c in t.columns if c.endswith('월') or '(당월)' in c],
                       height=420, spark='추이')

            # ② 전월 대비 증감
            st.markdown("---")
            st.subheader(f"② 전월 대비 증감 ({prev}월 → {base}월)")
            tot_m = df.groupby('월')['매출'].sum().reindex(months_all, fill_value=0)
            tp, tc = tot_m.get(prev, 0), tot_m.get(base, 0)
            a1, a2, a3 = st.columns(3)
            a1.metric(f"{prev}월 전체 매출", man(tp))
            a2.metric(f"{base}월 전체 매출", man(tc), delta=f"{(tc - tp) / tp * 100:+.1f}%" if tp else None)
            a3.metric("증가 / 감소 부대", f"{int(((cur - prv) > 0).sum())} / {int(((cur - prv) < 0).sum())}개")

            mom = (tot_m.pct_change() * 100).dropna().reset_index()
            mom.columns = ['월', '증감률']
            fig_mom = px.bar(mom, x='월', y='증감률', text=mom['증감률'].map(lambda v: f"{v:+.1f}%"),
                             color=np.where(mom['증감률'] >= 0, '상승', '하락'),
                             color_discrete_map={'상승': '#e5534b', '하락': '#3b8eea'})
            fig_mom.update_traces(textposition='outside', cliponaxis=False,
                                  marker_line_width=[3 if m == base else 0 for m in mom['월']], marker_line_color='white')
            fig_mom.update_layout(height=220, margin=dict(l=5, r=5, t=10, b=5), showlegend=False,
                                  yaxis=dict(title=None, showticklabels=False))
            fig_mom.update_xaxes(dtick=1, ticksuffix='월', title=None)
            st.plotly_chart(fig_mom, width='stretch')

            chg = pd.DataFrame({'부대명': cum.index, '등급': grade.values, f'{prev}월': prv.values, f'{base}월': cur.values})
            chg['증감액'] = chg[f'{base}월'] - chg[f'{prev}월']
            chg['증감률'] = np.where(chg[f'{prev}월'] > 0, chg['증감액'] / chg[f'{prev}월'].replace(0, np.nan) * 100, np.nan)
            flt = st.radio("보기", ["전체", "증가", "감소"], horizontal=True, key='chg_flt', label_visibility='collapsed')
            cv = chg if flt == "전체" else chg[chg['증감액'] > 0] if flt == "증가" else chg[chg['증감액'] < 0]
            show_table(cv.sort_values('증감액', ascending=False), num_cols=[f'{prev}월', f'{base}월', '증감액'],
                       pct_cols=['증감률'], height=360)

            # ③ 급증 / 급감
            st.markdown("---")
            SURGE = 30
            st.subheader(f"③ 매출 급증 / 급감 부대 (전월 대비 ±{SURGE}% 이상)")
            s = chg.copy()
            s.loc[(s[f'{prev}월'] == 0) & (s[f'{base}월'] > 0), '비고'] = '신규'
            s.loc[(s[f'{prev}월'] > 0) & (s[f'{base}월'] == 0), '비고'] = '발주 없음'
            up_df = s[(s['증감률'] >= SURGE) | (s['비고'] == '신규')].sort_values('증감액', ascending=False)
            dn_df = s[s['증감률'] <= -SURGE].sort_values('증감액')
            g1, g2 = st.columns(2)
            cols = ['부대명', '등급', f'{prev}월', f'{base}월', '증감률', '비고']
            with g1:
                st.markdown(f"**📈 급증 ({len(up_df)}개)**")
                show_table(up_df[cols], num_cols=[f'{prev}월', f'{base}월'], pct_cols=['증감률'])
            with g2:
                st.markdown(f"**📉 급감 ({len(dn_df)}개)**")
                show_table(dn_df[cols], num_cols=[f'{prev}월', f'{base}월'], pct_cols=['증감률'])

            # ④ ⑤ 꾸준히 증가 / 감소
            WIN = upto[-6:]
            st.markdown("---")
            if len(WIN) < 3:
                st.info("꾸준한 증가·감소 판단에는 3개월 이상의 데이터가 필요합니다.")
            else:
                rows = []
                x = np.arange(len(WIN))
                for u in cum.index:
                    y = up.loc[u, WIN].values.astype(float)
                    if y.mean() <= 0:
                        continue
                    slope, icpt = np.polyfit(x, y, 1)
                    fit = slope * x + icpt
                    ss = ((y - y.mean()) ** 2).sum()
                    r2 = 1 - ((y - fit) ** 2).sum() / ss if ss > 0 else 0
                    rows.append({'부대명': u, '등급': grade[u], **{f'{m}월': y[i] for i, m in enumerate(WIN)},
                                 '월평균 변화율': slope / y.mean() * 100, '_r2': r2, '추이': list(y)})
                tr = pd.DataFrame(rows)
                num = [f'{m}월' for m in WIN]
                cond_note = (f"{WIN[0]}월~{WIN[-1]}월({len(WIN)}개월) 추세선 기준 · 월평균 변화율 ±5% 이상이면서 "
                             "추세가 일정한(들쭉날쭉하지 않은) 부대")
                inc = tr[(tr['월평균 변화율'] >= 5) & (tr['_r2'] >= 0.4)].sort_values('월평균 변화율', ascending=False)
                dec = tr[(tr['월평균 변화율'] <= -5) & (tr['_r2'] >= 0.4)].sort_values('월평균 변화율')
                st.subheader(f"④ 매출이 꾸준히 증가한 부대 ({len(inc)}개)")
                st.caption(cond_note)
                if inc.empty:
                    st.caption("해당 부대 없음")
                else:
                    show_table(inc[['부대명', '등급'] + num + ['월평균 변화율', '추이']], num_cols=num,
                               pct_cols=['월평균 변화율'], spark='추이')
                st.markdown("---")
                st.subheader(f"⑤ 매출이 꾸준히 감소한 부대 ({len(dec)}개)")
                st.caption(cond_note)
                if dec.empty:
                    st.caption("해당 부대 없음")
                else:
                    show_table(dec[['부대명', '등급'] + num + ['월평균 변화율', '추이']], num_cols=num,
                               pct_cols=['월평균 변화율'], spark='추이')

    except Exception as e:
        import traceback
        st.error(f"🚨 오류 발생: {e}")
        st.code(traceback.format_exc())
