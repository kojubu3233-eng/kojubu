import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import numpy as np
from auto.units import UNIT_MASTER, MIN_DELIVERIES, FOCUS_UNITS, normalize_unit, delivery_key, closed_through_month
import plotly.io as pio
from datetime import datetime, timedelta, timezone

def safe_sorted(iterable):
    return sorted([str(x) for x in iterable if str(x) not in ('nan', 'None', '')])

st.set_page_config(page_title="연간 매출 분석 대시보드", layout="wide", initial_sidebar_state="expanded")

# ── 디자인 (Apps Script 영업계획 대시보드 스타일) ──
ACCENT, ACCENT_DIM, RED, BLUE = "#6c8ef5", "#3d4a72", "#f06c6c", "#5b9cf5"
pio.templates["kojubu"] = go.layout.Template(layout=dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#c9cdd6", size=12), colorway=[ACCENT, "#f0a35c", "#4cc38a", "#e05fa0", "#9b7cf0", "#3cc0d0", "#c9a13c"],
    xaxis=dict(gridcolor="#232836", zerolinecolor="#232836"), yaxis=dict(gridcolor="#232836", zerolinecolor="#232836"),
    hoverlabel=dict(bgcolor="#1b2030")))
pio.templates.default = "kojubu"
px.defaults.template = "kojubu"

st.markdown("""
<style>
.block-container {padding-top: 2.2rem; max-width: 1500px;}
h1 {font-size: 1.55rem !important; font-weight: 700 !important;}
[data-testid="stVerticalBlockBorderWrapper"], div[data-testid="stVerticalBlock"][data-border="true"]
    {background:#161a23; border:1px solid #2a2f3a !important; border-radius:12px;}
.stTabs [data-baseweb="tab-list"] {gap:4px; background:#161a23; padding:4px; border-radius:10px; width:fit-content;}
.stTabs [data-baseweb="tab"] {padding:6px 18px; border-radius:8px; height:auto;}
.stTabs [aria-selected="true"] {background:#262c3a;}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {display:none;}
.kp {background:#161a23;border:1px solid #2a2f3a;border-radius:12px;padding:14px 16px;height:100%;}
.kp .l {font-size:12.5px;color:#9aa1ae} .kp .v {font-size:26px;font-weight:700;margin:4px 0 2px}
.kp .s {font-size:12px;color:#7d8594}
.sec {font-size:16px;font-weight:700;margin:2px 0 2px} .sec-sub {font-size:12.5px;color:#8b93a1;margin-bottom:6px}
.sumgrid {display:grid;grid-template-columns:repeat(auto-fit,minmax(330px,1fr));gap:12px;}
.sc {background:#161a23;border:1px solid #2a2f3a;border-radius:12px;padding:14px 16px;}
.sc h4 {font-size:14.5px;margin:0 0 2px;font-weight:700;} .sc .d {font-size:12px;color:#8b93a1;margin-bottom:8px}
.sc table {width:100%;border-collapse:collapse;font-size:12.5px}
.sc th {color:#8b93a1;font-weight:500;text-align:right;padding:5px 4px;border-bottom:1px solid #2a2f3a}
.sc th:first-child, .sc td:first-child {text-align:left}
.sc td {padding:6px 4px;border-bottom:1px solid #20242f;text-align:right;white-space:nowrap}
.sc .neg {color:#f06c6c} .sc .pos {color:#4cc38a} .sc .none {color:#8b93a1;text-align:center;padding:14px}
.tag {display:inline-block;font-size:10.5px;padding:1px 6px;border-radius:6px;background:#26314f;color:#8fb0ff;margin-left:4px}
.dot {display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:5px;vertical-align:middle}
.foot {font-size:12px;color:#7d8594;margin-top:8px}
</style>""", unsafe_allow_html=True)


def kpi(label, value, sub="", color="#e6e8ee"):
    return f"<div class='kp'><div class='l'>{label}</div><div class='v' style='color:{color}'>{value}</div><div class='s'>{sub}</div></div>"


def sec(title, sub=""):
    st.markdown(f"<div class='sec'>{title}</div>" + (f"<div class='sec-sub'>{sub}</div>" if sub else ""),
                unsafe_allow_html=True)


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
            d = d.copy()
            for c in list(num_cols) + list(pct_cols):
                d[c] = pd.to_numeric(d[c], errors='coerce')   # 빈 값이 'None'으로 보이지 않도록 숫자형 강제
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
            best_m = int(monthly.idxmax())
            k1, k2, k3, k4 = st.columns(4)
            k1.markdown(kpi("올해 누적 매출", man(total), f"{total:,.0f}원", ACCENT), unsafe_allow_html=True)
            k2.markdown(kpi(f"{cur_m}월 매출", man(monthly.get(cur_m, 0)),
                            "확정+예정 (미마감)" if cur_m > _closed else "확정 (마감)"), unsafe_allow_html=True)
            k3.markdown(kpi("월평균 매출", man(monthly[monthly.index <= cur_m].mean()), f"1~{cur_m}월 기준"), unsafe_allow_html=True)
            k4.markdown(kpi("최고 월", f"{best_m}월", f"{monthly.max():,.0f}원"), unsafe_allow_html=True)
            st.write("")

            mdf = monthly.reset_index()
            mdf.columns = ['월', '매출']
            mdf['구분'] = np.where(mdf['월'] <= _closed, '확정(마감)', '확정+예정(미마감)')
            fig_m = px.bar(mdf, x='월', y='매출', color='구분', text=mdf['매출'].map(man),
                           color_discrete_map={'확정(마감)': ACCENT, '확정+예정(미마감)': ACCENT_DIM})
            fig_m.update_traces(textposition='outside', cliponaxis=False)
            fig_m.update_layout(height=300, margin=dict(l=10, r=10, t=30, b=10), yaxis_title=None,
                                legend=dict(orientation='h', y=1.12, x=0, title=None))
            fig_m.update_xaxes(dtick=1, ticksuffix='월', title=None)
            fig_m.update_yaxes(showticklabels=False)
            with st.container(border=True):
                sec("월별 매출", "연한 막대 = 미마감 월(확정+예정)")
                st.plotly_chart(fig_m, width='stretch')

            # ── 품목 순위 (상위 10 / 하위 10 한 줄) ──
            with st.container(border=True):
                sec("🏆 품목 매출 순위")
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
                    item_chart(items.nlargest(10, '매출').sort_values('매출'), RED, 'top10', '🔴 상위 10개 품목')
                with cR:
                    item_chart(items.nsmallest(10, '매출').sort_values('매출', ascending=False), BLUE, 'bottom10', '🔵 하위 10개 품목')

            # ── 구분별 카드 ──
            with st.container(border=True):
                sec("🗂️ 구분별 매출")
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
                                index=next((i for i, m in enumerate(months_desc) if m <= TARGET_M), 0),
                                format_func=lambda x: f"{x}월")
            prev = base - 1
            t2.markdown(f"<div style='padding-top:34px;font-size:12.5px;color:#8b93a1'>"
                        f"표 행 색상: <span style='background:rgba(229,83,75,.3);padding:1px 6px;border-radius:4px'>상위</span> "
                        f"누적 매출 상위 1/3 · <span style='background:rgba(59,142,234,.3);padding:1px 6px;border-radius:4px'>하위</span> "
                        f"하위 1/3 (1월\\~{base}월 누적)</div>", unsafe_allow_html=True)

            upto = [m for m in months_all if m <= base]
            up = unit_pivot_raw[upto]
            cum = up.sum(axis=1)
            if len(cum) >= 3:
                grade = pd.qcut(cum.rank(method='first'), q=3, labels=[LO, MID, HI]).astype(str)
            else:
                grade = pd.Series(MID, index=cum.index)
            cur = unit_pivot_raw.get(base, pd.Series(0, index=unit_pivot_raw.index))
            prv = unit_pivot_raw.get(prev, pd.Series(0, index=unit_pivot_raw.index))

            # ── 분석 요약 (영업활동 수립용 한눈 정리) ──
            def _v(u, m):
                return float(unit_pivot_raw.at[u, m]) if m in unit_pivot_raw.columns else 0.0


            def _name(u):
                g = grade.get(u, MID)
                c = '#f06c6c' if g == HI else '#5b9cf5' if g == LO else '#9aa1ae'
                return f"<span class='dot' style='background:{c}'></span>{u}" + ("<span class='tag'>중점</span>" if u in FOCUS_UNITS else "")


            def _pct(a, b):
                if a <= 0:
                    return "<span class='pos'>신규</span>" if b > 0 else "-"
                r = (b - a) / a * 100
                return f"<span class='{'neg' if r < 0 else 'pos'}'>{r:+.1f}%</span>"


            def _card(title, desc, head, rows, empty="해당 부대 없음"):
                th = "".join(f"<th>{h}</th>" for h in head)
                body = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in rows) \
                    or f"<tr><td class='none' colspan='{len(head)}'>{empty}</td></tr>"
                return f"<div class='sc'><h4>{title}</h4><div class='d'>{desc}</div><table><tr>{th}</tr>{body}</table></div>"


            units_all = list(unit_pivot_raw.index)
            m3 = [m for m in (prev - 2, prev - 1, prev) if m in unit_pivot_raw.columns]
            cards = []
            if len(m3) == 3:
                a_, b_, c_ = m3
                seq = [(u, _v(u, a_), _v(u, b_), _v(u, c_)) for u in units_all]
                down3 = sorted([x for x in seq if x[1] > x[2] > x[3] and x[1] > 0], key=lambda x: (x[3] - x[1]) / x[1])[:7]
                up3 = sorted([x for x in seq if 0 < x[1] < x[2] < x[3]], key=lambda x: -(x[3] - x[1]) / x[1])[:7]
                cards.append(_card("📉 3개월 연속 하락 부대", f"{a_}월 → {b_}월 → {c_}월 매출이 계속 감소 · 우선 방문 대상",
                                   ["부대", "매출 추이", "증감"],
                                   [[_name(u), f"{man(x)} → {man(y)} → {man(z)}", _pct(x, z)] for u, x, y, z in down3]))
                cards.append(_card("📈 3개월 연속 상승 부대", f"{a_}월 → {b_}월 → {c_}월 매출이 계속 증가 · 품목 확대 제안 대상",
                                   ["부대", "매출 추이", "증감"],
                                   [[_name(u), f"{man(x)} → {man(y)} → {man(z)}", _pct(x, z)] for u, x, y, z in up3]))
            if prev - 1 in unit_pivot_raw.columns and prev in unit_pivot_raw.columns:
                dd_ = sorted([(u, _v(u, prev - 1), _v(u, prev)) for u in units_all], key=lambda x: x[2] - x[1])
                dd_ = [x for x in dd_ if x[2] < x[1]][:7]
                cards.append(_card("🔻 전월 대비 감소 상위", f"{prev - 1}월 → {prev}월 감소액이 큰 순",
                                   ["부대", "매출", "감소"],
                                   [[_name(u), f"{man(x)} → {man(y)}", f"<span class='neg'>-{man(x - y)}</span> ({_pct(x, y)})"] for u, x, y in dd_]))
            nx = [(u, _v(u, prev), _v(u, base)) for u in units_all]
            nx = sorted([x for x in nx if x[1] > 0 and x[2] < x[1] * 0.5], key=lambda x: -x[1])[:8]
            cards.append(_card(f"⚠️ {base}월 {'예정' if base > _closed else '매출'} 없음·급감",
                               f"{prev}월 매출은 있으나 {base}월 {'예정' if base > _closed else '매출'}이 없거나 50% 미만 · 발주 확인 필요",
                               ["부대", f"{prev}월", f"{base}월"],
                               [[_name(u), man(x), "<span class='neg'>없음</span>" if y == 0 else f"{man(y)} ({_pct(x, y)})"] for u, x, y in nx]))
            fr = [u for u in FOCUS_UNITS if u in unit_pivot_raw.index]
            cards.append(_card("🎯 중점 부대 현황", f"{prev}월 → {base}월 · 누적 기준 등급 색상", ["부대", f"{prev}월", f"{base}월", "증감"],
                               [[_name(u), man(_v(u, prev)), man(_v(u, base)), _pct(_v(u, prev), _v(u, base))] for u in fr]))
            sec("📋 분석 요약", f"{base}월 기준 영업활동 우선순위")
            st.markdown("<div class='sumgrid'>" + "".join(cards) + "</div>"
                        f"<div class='foot'>금액은 마감 월(~{_closed}월)은 확정, 이후 월은 확정+예정 기준 · "
                        "<span class='dot' style='background:#f06c6c'></span>상위 1/3 "
                        "<span class='dot' style='background:#5b9cf5'></span>하위 1/3 · <span class='tag'>중점</span> 중점 부대</div>",
                        unsafe_allow_html=True)
            st.write("")

            # ① 부대별 매출 현황
            with st.container(border=True):
                sec(f"① 부대별 매출 현황 ({base}월 기준)")
                t = pd.DataFrame({'부대명': cum.index, '등급': grade.values, '누적매출': cum.values,
                                  '월평균': (cum / max(len(upto), 1)).values})
                pcol, ccol = f"{prev}월(전월)", f"{base}월(당월)"
                t[pcol] = prv.reindex(cum.index).fillna(0).values
                t[ccol] = cur.reindex(cum.index).fillna(0).values
                t['증감률'] = pd.to_numeric(np.where(t[pcol] > 0, (t[ccol] - t[pcol]) / t[pcol].replace(0, np.nan) * 100, np.nan), errors='coerce')
                t['변동'] = np.select([t['증감률'] >= 50, t['증감률'] <= -50, (t[pcol] == 0) & (t[ccol] > 0)],
                                    ['📈 급증', '📉 급감', '🆕 신규'], default='')
                t['추이'] = [list(up.loc[u].values) for u in cum.index]
                t = t[t['누적매출'] > 0].sort_values('누적매출', ascending=False)
                st.caption("변동: 전월 대비 ±50% 이상이면 📈 급증 / 📉 급감 표시")
                t = t[['부대명', '등급', '누적매출', '월평균', pcol, ccol, '증감률', '변동', '추이']]
                show_table(t, num_cols=['누적매출', '월평균', pcol, ccol], pct_cols=['증감률'],
                           height=420, spark='추이')

            # ② 당월 매출 급증 / 급감 (전월 대비 ±50%)
            with st.container(border=True):
                sec(f"② {base}월 매출 급증 / 급감 부대")
                st.caption(f"{prev}월(전월) 대비 {base}월(당월) 매출이 ±50% 이상 변한 부대"
                           + (" · 미마감 월은 예정 포함" if base > _closed else ""))
                sv = t.copy()
                sv['증감액'] = sv[ccol] - sv[pcol]
                surge_up = sv[sv['변동'].isin(['📈 급증', '🆕 신규'])].sort_values('증감액', ascending=False)
                surge_dn = sv[sv['변동'] == '📉 급감'].sort_values('증감액')
                scols = ['부대명', '등급', pcol, ccol, '증감액', '증감률', '변동']
                s1, s2 = st.columns(2)
                with s1:
                    st.markdown(f"**📈 급증 ({len(surge_up)}개)**")
                    show_table(surge_up[scols], num_cols=[pcol, ccol, '증감액'], pct_cols=['증감률'])
                with s2:
                    st.markdown(f"**📉 급감 ({len(surge_dn)}개)**")
                    show_table(surge_dn[scols], num_cols=[pcol, ccol, '증감액'], pct_cols=['증감률'])

            # ③ 분기 분석
            with st.container(border=True):
                sec("③ 분기 분석")
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
                qd[f'{qsel}분기 합계'] = unit_pivot_raw[qms].sum(axis=1).values
                qnum = [f'{m}월' for m in qms] + [f'{qsel}분기 합계']
                if pms:
                    qd[f'{qsel - 1}분기 합계'] = unit_pivot_raw[pms].sum(axis=1).values
                    qd['증감률'] = np.where(qd[f'{qsel - 1}분기 합계'] > 0,
                                         (qd[f'{qsel}분기 합계'] - qd[f'{qsel - 1}분기 합계']) / qd[f'{qsel - 1}분기 합계'].replace(0, np.nan) * 100, np.nan)
                    qcols = ['부대명', '등급', f'{qsel - 1}분기 합계'] + qnum + ['증감률']
                    rise = qd[qd['증감률'] >= QTH].sort_values('증감률', ascending=False)
                    fall = qd[qd['증감률'] <= -QTH].sort_values('증감률')
                    r1, r2 = st.columns(2)
                    with r1:
                        st.markdown(f"**📈 상승 부대 ({len(rise)}개)** · 전 분기 대비 +{QTH}% 이상")
                        show_table(rise[qcols], num_cols=[f'{qsel - 1}분기 합계'] + qnum, pct_cols=['증감률'])
                    with r2:
                        st.markdown(f"**📉 하락 부대 ({len(fall)}개)** · 전 분기 대비 −{QTH}% 이상")
                        show_table(fall[qcols], num_cols=[f'{qsel - 1}분기 합계'] + qnum, pct_cols=['증감률'])
                qavg = qd[[f'{m}월' for m in qms]].mean(axis=1)
                qd['등락폭'] = np.where(qavg > 0, (qd[[f'{m}월' for m in qms]].max(axis=1) - qd[[f'{m}월' for m in qms]].min(axis=1)) / qavg.replace(0, np.nan) * 100, np.nan)
                swing = qd[(qd['등락폭'] >= SWING) & (qd[f'{qsel}분기 합계'] > 0)].sort_values('등락폭', ascending=False)
                st.markdown(f"**🌊 등락폭이 큰 부대 ({len(swing)}개)** · 분기 내 (최고월 − 최저월) ÷ 월평균 ≥ {SWING}% → 발주가 불규칙한 부대")
                show_table(swing[['부대명', '등급'] + qnum + ['등락폭']], num_cols=qnum, pct_cols=['등락폭'])

    except Exception as e:
        import traceback
        st.error(f"🚨 오류 발생: {e}")
        st.code(traceback.format_exc())
