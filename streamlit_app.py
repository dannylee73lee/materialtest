import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os, io

st.set_page_config(page_title="자재 현황 대시보드", page_icon="📦", layout="wide")

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #FAFAF8; }
[data-testid="stSidebar"]          { background: #F5F3EE; border-right: 1px solid #E8E4DC; }
.kpi-card {
    background: #FFFFFF; border: 0.5px solid #E8E4DC;
    border-radius: 12px; padding: 20px 24px; text-align: center;
}
.kpi-label { color: #7C7268; font-size: 13px; margin-bottom: 6px; }
.kpi-value { color: #1A1A1A; font-size: 32px; font-weight: 600; }
.kpi-unit  { color: #A89E94; font-size: 13px; margin-top: 4px; }
.section-title {
    color: #3D3530; font-size: 15px; font-weight: 600;
    padding: 8px 0 5px 0; border-bottom: 2px solid #B45309;
    margin-bottom: 12px; display: inline-block;
}
.status-ok   { color: #166634; font-weight: 600; }
.status-warn { color: #B45309; font-weight: 600; }
[data-testid="stTabs"] button {
    font-size: 14px !important; font-weight: 600 !important; color: #3D3530 !important;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    border-bottom-color: #B45309 !important; color: #1A1A1A !important;
}
</style>
""", unsafe_allow_html=True)

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
MAPPING_PATH = os.path.join(BASE_DIR, "data", "결과_format.xlsx")

# 탭별 설정: sheet=시트명, col_map=컬럼명 통일(세분류→소분류, 제조업체→제조사), filter2=2차필터
TAB_CONFIG = [
    {
        "label":   "📡 5G 물자",
        "sheet":   "5G물자",
        "col_map": {
            "대분류":"대분류","중분류":"중분류","소분류":"소분류",
            "자재코드":"자재코드","품명":"품명","제조사":"제조사",
            "이월(신품)":"이월(신품)","이월(구품)":"이월(구품)","이월(재고)":"이월(재고)",
            "신품":"신품","구품":"구품","재고":"재고","전체":"전체"
        },
        "filter2": "중분류",
    },
    {
        "label":   "📶 5G·LTE 물자",
        "sheet":   "5G,LTE물자",
        "col_map": {
            "대분류":"대분류","소분류":"소분류",
            "자재코드":"자재코드","품명":"품명","제조사":"제조사",
            "이월(신품)":"이월(신품)","이월(구품)":"이월(구품)","이월(재고)":"이월(재고)",
            "신품":"신품","구품":"구품","재고":"재고","전체":"전체"
        },
        "filter2": "소분류",
    },
    {
        "label":   "🔧 RRU·MiBos·W·설비물자",
        "sheet":   "RRU, MiBos, W, 설비물자",
        "col_map": {
            "대분류":"대분류","중분류":"중분류",
            "세분류":"소분류",
            "자재코드":"자재코드","품명":"품명",
            "제조업체":"제조사",
            "이월(신품)":"이월(신품)","이월(구품)":"이월(구품)","이월(재고)":"이월(재고)",
            "신품":"신품","구품":"구품","재고":"재고","전체":"전체"
        },
        "filter2": "중분류",
    },
    {
        "label":   "🏢 5G·LTE(인빌딩·Legacy)",
        "sheet":   "5G, LTE(인빌딩, Legacy)",
        "col_map": {
            "구분":"대분류","구분1":"소분류",
            "자재코드":"자재코드","품명":"품명","제조사":"제조사",
            "이월(신품)":"이월(신품)","이월(구품)":"이월(구품)","이월(재고)":"이월(재고)",
            "신품":"신품","구품":"구품","재고":"재고","전체":"전체"
        },
        "filter2": "소분류",
    },
    {
        "label":   "📡 QMHS·IMHS·AMHS",
        "sheet":   "QMHS, IMHS, AMHS",
        "col_map": {
            "구분":"대분류",
            "자재코드":"자재코드","품명":"품명","제조사":"제조사",
            "이월(신품)":"이월(신품)","이월(구품)":"이월(구품)","이월(재고)":"이월(재고)",
            "신품":"신품","구품":"구품","재고":"재고","전체":"전체"
        },
        "filter2": "대분류",
    },
    {
        "label":   "🔌 SFP·MUX·OCM",
        "sheet":   "SFP, MUX, OCM",
        "col_map": {
            "대분류":"대분류","중분류":"중분류",
            "자재코드":"자재코드","자재명":"품명",
            "이월(신품)":"이월(신품)","이월(구품)":"이월(구품)","이월(재고)":"이월(재고)",
            "신품":"신품","구품":"구품","재고":"재고","전체":"전체"
        },
        "filter2": "중분류",
    },
    {
        "label":   "📦 부대물자",
        "sheet":   "부대물자",
        "col_map": {
            "분류":"대분류",
            "자재코드":"자재코드","품명":"품명",
            "이월(신품)":"이월(신품)","이월(구품)":"이월(구품)","이월(재고)":"이월(재고)",
            "신품":"신품","구품":"구품","재고":"재고","전체":"전체"
        },
        "filter2": "대분류",
    },
    {
        "label":   "📡 안테나",
        "sheet":   "안테나",
        "col_map": {
            "자재코드":"자재코드","품명":"품명",
            "이월(신품)":"이월(신품)","이월(구품)":"이월(구품)","이월(재고)":"이월(재고)",
            "신품":"신품","구품":"구품","재고":"재고","전체":"전체"
        },
        "filter2": None,
    },
]

def fmt(val):
    return '-' if val == 0 else f'{val:,}'

@st.cache_data
def load_mapping_sheet(path, sheet, col_map):
    try:
        df = pd.read_excel(path, sheet_name=sheet, header=0)
        df = df.rename(columns=col_map)
        df['자재코드'] = pd.to_numeric(df['자재코드'], errors='coerce').astype('Int64')
        df = df.dropna(subset=['자재코드'])
        for c in ['신품','구품','재고']:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)
        return df
    except Exception as e:
        st.error(f"매핑 시트 [{sheet}] 읽기 실패: {e}")
        return pd.DataFrame()

def parse_excel_carryover(file_bytes, file_name):
    """이월현황 파싱 (컬럼 10개, 자재분류/자재명 없음)"""
    buf = io.BytesIO(file_bytes)
    raw = pd.read_excel(buf, header=None)
    header_row = 0
    for i in range(min(5, len(raw))):
        row_vals = raw.iloc[i].astype(str).tolist()
        if any('순번' in v or '자재코드' in v for v in row_vals):
            header_row = i
    buf.seek(0)
    df = pd.read_excel(buf, header=header_row)
    if len(df.columns) < 10:
        raise ValueError(f"컬럼 수 {len(df.columns)}개 — 10개 이상 필요")
    df = df.iloc[:, :10]
    df.columns = ['순번','사업년도','지역본부','군','업체명',
                  '자재코드','FULL자재명','신품','구품_양호','구품_불량']
    df['자재분류'] = '이월'
    df['자재명']   = df['FULL자재명']
    df = df[pd.to_numeric(df['순번'], errors='coerce').notna()]
    df['신품']     = pd.to_numeric(df['신품'],      errors='coerce').fillna(0).astype(int)
    df['구품']     = pd.to_numeric(df['구품_양호'], errors='coerce').fillna(0).astype(int)
    df['재고']     = df['신품'] + df['구품']
    df['자재코드'] = pd.to_numeric(df['자재코드'],  errors='coerce').astype('Int64')
    df['파일명']   = file_name
    return df.dropna(subset=['자재코드'])

def parse_excel(file_bytes, file_name):
    buf = io.BytesIO(file_bytes)
    raw = pd.read_excel(buf, header=None)
    # 헤더 행 탐색: '순번'이 있는 마지막 행 (2중 헤더 대응)
    header_row = 0
    for i in range(min(5, len(raw))):
        row_vals = raw.iloc[i].astype(str).tolist()
        if any('순번' in v or '자재코드' in v for v in row_vals):
            header_row = i
    buf.seek(0)
    df = pd.read_excel(buf, header=header_row)
    if len(df.columns) < 12:
        raise ValueError(f"컬럼 수 {len(df.columns)}개 — 12개 이상 필요")
    df = df.iloc[:, :12]
    df.columns = ['순번','사업년도','지역본부','군','업체명','자재분류',
                  '자재코드','자재명','FULL자재명','신품','구품_양호','구품_불량']
    # 순번이 숫자인 행만 유지 (헤더 중복 행 제거)
    df = df[pd.to_numeric(df['순번'], errors='coerce').notna()]
    df['신품']     = pd.to_numeric(df['신품'],      errors='coerce').fillna(0).astype(int)
    df['구품']     = pd.to_numeric(df['구품_양호'], errors='coerce').fillna(0).astype(int)
    df['재고']     = df['신품'] + df['구품']
    df['자재코드'] = pd.to_numeric(df['자재코드'],  errors='coerce').astype('Int64')
    df['파일명']   = file_name
    return df.dropna(subset=['자재코드'])

# ── 사이드바 ─────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("#### 📂 수불부 - 현재")
    st.caption("2026년 재고현황")
    uploaded_current = st.file_uploader(
        "현재 수불부", type=["xlsx","xls"],
        accept_multiple_files=True, label_visibility="collapsed",
        key="upload_current"
    )
    st.markdown("---")
    st.markdown("#### 📂 수불부 - 이월 현황")
    st.caption("2025년 이월 재고현황")
    uploaded_carry = st.file_uploader(
        "이월 수불부", type=["xlsx","xls"],
        accept_multiple_files=True, label_visibility="collapsed",
        key="upload_carry"
    )
    uploaded_files = (uploaded_current or []) + (uploaded_carry or [])

# ── 메인 헤더 ────────────────────────────────────────────────────
st.markdown("# 📦 자재 현황 대시보드")

if not uploaded_files:
    st.markdown("---")
    st.info("**수불부 파일을 업로드하세요**\n\n좌측 사이드바에서 수불부 현재 또는 이월내역 파일을 선택하면 대시보드가 생성됩니다.")
    st.stop()

# ── 수불부 파싱 ──────────────────────────────────────────────────
all_dfs, parse_errors = [], []

for f in (uploaded_current or []):
    try:
        parsed = parse_excel(f.read(), f.name)
        if parsed is not None and len(parsed) > 0:
            all_dfs.append(parsed)
        else:
            parse_errors.append(f"{f.name}: 데이터 없음")
    except Exception as e:
        parse_errors.append(f"{f.name}: {e}")

for f in (uploaded_carry or []):
    try:
        parsed = parse_excel_carryover(f.read(), f.name)
        if parsed is not None and len(parsed) > 0:
            all_dfs.append(parsed)
        else:
            parse_errors.append(f"{f.name}: 데이터 없음")
    except Exception as e:
        parse_errors.append(f"{f.name}: {e}")

if parse_errors:
    for msg in parse_errors:
        st.error(f"❌ 파싱 오류: {msg}")
        print(f"[PARSE ERROR] {msg}")  # 터미널 출력

if not all_dfs:
    st.error("업로드된 파일에서 데이터를 읽지 못했습니다.")
else:

    df_raw = pd.concat(all_dfs, ignore_index=True)
    # 현재 수불부: 신품/구품/재고
    current_raw = df_raw[df_raw['자재분류'] != '이월'] if '자재분류' in df_raw.columns else df_raw
    carry_raw   = df_raw[df_raw['자재분류'] == '이월'] if '자재분류' in df_raw.columns else pd.DataFrame(columns=df_raw.columns)

    qty_current = current_raw.groupby('자재코드')[['신품','구품','재고']].sum().reset_index()
    qty_carry   = carry_raw.groupby('자재코드')[['신품','구품','재고']].sum().reset_index()
    qty_carry   = qty_carry.rename(columns={'신품':'이월(신품)','구품':'이월(구품)','재고':'이월(재고)'})

    qty_df = pd.merge(qty_current, qty_carry, on='자재코드', how='outer').fillna(0)
    for c in ['신품','구품','재고','이월(신품)','이월(구품)','이월(재고)']:
        qty_df[c] = qty_df[c].astype(int)
    qty_df['전체'] = qty_df['재고'] + qty_df['이월(재고)']

    file_tag = " | ".join(f"📄 {f}" for f in df_raw['파일명'].unique())
    st.markdown(f"<small style='color:#A89E94'>{file_tag}</small>", unsafe_allow_html=True)
    st.markdown("---")

    # ── 전체 KPI ─────────────────────────────────────────────────────
    total_rows  = len(df_raw)
    total_new   = int(df_raw['신품'].sum())
    total_used  = int(df_raw['구품'].sum())
    total_stock = int(df_raw['재고'].sum())
    has_qty_cnt = int((df_raw['재고'] > 0).sum())

    def kpi(col, label, value, unit=""):
        col.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value:,}</div>
            <div class="kpi-unit">{unit}</div>
        </div>""", unsafe_allow_html=True)

    k1, k2, k3, k4, k5 = st.columns(5)
    kpi(k1, "전체 자재 항목", total_rows,    "건")
    kpi(k2, "신품 수량",      total_new,     "개")
    kpi(k3, "구품 수량",      total_used,    "개")
    kpi(k4, "총 재고",        total_stock,   "개")
    kpi(k5, "재고 보유 항목", has_qty_cnt,   "건")
    st.markdown("<br>", unsafe_allow_html=True)

    # ── 차트 ─────────────────────────────────────────────────────────
    LAYOUT       = dict(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color='#3D3530')
    MARGIN_SM    = dict(t=20, b=20)
    MARGIN_LABEL = dict(t=30, b=20)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-title">📊 신품 / 구품 비율</div>', unsafe_allow_html=True)
        if total_stock > 0:
            pie_df = pd.DataFrame({'구분': ['신품','구품'], '수량': [total_new, total_used]})
            fig = px.pie(pie_df, names='구분', values='수량',
                         color_discrete_sequence=['#5B9BD5','#F59E0B'], hole=0.45)
            fig.update_traces(textinfo='label+percent', textfont_size=13)
            fig.update_layout(**LAYOUT, margin=MARGIN_SM, legend=dict(font=dict(color='#3D3530')))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("수량 데이터가 없습니다.")

    with c2:
        st.markdown('<div class="section-title">🏢 업체별 재고 현황 (TOP 10)</div>', unsafe_allow_html=True)
        biz = df_raw[df_raw['재고'] > 0].groupby('업체명')[['신품','구품']].sum().reset_index()
        biz['재고'] = biz['신품'] + biz['구품']
        biz = biz.sort_values('재고', ascending=False).head(10)
        if not biz.empty:
            total_biz = biz['재고'].sum()
            biz['비율'] = (biz['재고'] / total_biz * 100).round(1)
            biz['label'] = biz['비율'].apply(lambda x: f'{x}%')
            fig = go.Figure()
            fig.add_bar(x=biz['업체명'], y=biz['신품'], name='신품', marker_color='#5B9BD5')
            fig.add_bar(x=biz['업체명'], y=biz['구품'], name='구품', marker_color='#F59E0B',
                        text=biz['label'], textposition='outside', textfont=dict(size=11))
            fig.update_layout(barmode='stack', **LAYOUT, margin=dict(t=30,b=20),
                              xaxis=dict(gridcolor='#E8E4DC'), yaxis=dict(gridcolor='#E8E4DC'),
                              legend=dict(font=dict(color='#3D3530')))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("업체 데이터가 없습니다.")

    with st.container():
        st.markdown('<div class="section-title">📦 재고 TOP 10 자재</div>', unsafe_allow_html=True)
        top_base = (df_raw[df_raw['재고'] > 0].groupby('자재명')[['신품','구품','재고']].sum()
                    .reset_index().sort_values('재고', ascending=True).tail(10))
        if not top_base.empty:
            top_base['자재명_short'] = top_base['자재명'].str[:25]
            fig = go.Figure()
            fig.add_bar(y=top_base['자재명_short'], x=top_base['신품'], name='신품',
                        orientation='h', marker_color='#5B9BD5',
                        text=top_base['신품'].apply(lambda v: f'{v:,}' if v > 0 else ''),
                        textposition='inside', insidetextanchor='middle')
            fig.add_bar(y=top_base['자재명_short'], x=top_base['구품'], name='구품',
                        orientation='h', marker_color='#F59E0B',
                        text=top_base['재고'].apply(lambda v: f'{v:,}'),
                        textposition='outside')
            fig.update_layout(barmode='stack', **LAYOUT, margin=dict(t=20,b=20,r=70),
                              xaxis=dict(gridcolor='#E8E4DC'), yaxis=dict(gridcolor='#E8E4DC'),
                              legend=dict(font=dict(color='#3D3530')))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("재고 데이터가 없습니다.")



    st.markdown("---")

    # ── 탭별 상세 테이블 ─────────────────────────────────────────────
    if not os.path.exists(MAPPING_PATH):
        st.warning("data/결과_format.xlsx 가 없어 상세 테이블을 표시할 수 없습니다.")
        st.stop()

    tabs = st.tabs([t["label"] for t in TAB_CONFIG])

    for tab_ui, tab_cfg in zip(tabs, TAB_CONFIG):
        with tab_ui:
            mp = load_mapping_sheet(MAPPING_PATH, tab_cfg["sheet"], tab_cfg["col_map"])
            if mp.empty:
                st.warning(f"[{tab_cfg['sheet']}] 시트를 읽을 수 없습니다.")
                continue

            # 수불부 수량 LEFT JOIN (매핑 기준)
            merged = pd.merge(mp, qty_df, on='자재코드', how='left', suffixes=('_mp','_qty'))
            for c in ['신품','구품','재고','이월(신품)','이월(구품)','이월(재고)','전체']:
                qty_col = f'{c}_qty' if f'{c}_qty' in merged.columns else c
                merged[c] = pd.to_numeric(merged.get(qty_col, 0), errors='coerce').fillna(0).astype(int)
                for drop_c in [f'{c}_mp', f'{c}_qty']:
                    if drop_c in merged.columns:
                        merged = merged.drop(columns=[drop_c])

            tab_key = tab_cfg["label"]
            filter2 = tab_cfg["filter2"]
            # 표시 컬럼 = col_map 의 value(통일된 이름) 순서
            display_cols = list(tab_cfg["col_map"].values())

            # ── 필터 UI ──────────────────────────────────────────────
            st.markdown(
                "<div style='background:#FEF3C7;border:0.5px solid #F59E0B;border-radius:10px;"
                "padding:14px 18px;margin-bottom:14px'>",
                unsafe_allow_html=True
            )
            has_c1 = '대분류' in merged.columns
            if filter2 is not None and has_c1:
                fa, fb, fc, fd = st.columns([3, 2, 2, 1])
            elif has_c1:
                fa, fb, fd = st.columns([3, 2, 1])
                fc = None
            else:
                fa, fd = st.columns([5, 1])
                fb = fc = None

            with fa:
                kw = st.text_input("🔍 품명 검색", placeholder="품명 키워드를 입력하세요",
                                   key=f"kw_{tab_key}")
            sel_c1 = '전체'
            if fb is not None:
                with fb:
                    c1_opts = ['전체'] + sorted(merged['대분류'].dropna().unique().tolist())
                    sel_c1  = st.selectbox("대분류", c1_opts, key=f"c1_{tab_key}")
            sel_c2 = '전체'
            if fc is not None and filter2 is not None:
                with fc:
                    src2    = merged if sel_c1 == '전체' else merged[merged['대분류'] == sel_c1]
                    c2_opts = ['전체'] + sorted(src2[filter2].dropna().unique().tolist())
                    sel_c2  = st.selectbox(filter2, c2_opts, key=f"c2_{tab_key}")
            with fd:
                st.markdown("<div style='padding-top:28px'></div>", unsafe_allow_html=True)
                only_qty = st.checkbox("재고만", value=False, key=f"qty_{tab_key}")
            st.markdown("</div>", unsafe_allow_html=True)

            # ── 필터 적용 ────────────────────────────────────────────
            tdf = merged.copy()
            if kw:
                tdf = tdf[tdf['품명'].str.contains(kw, na=False, case=False)]
            if sel_c1 != '전체' and '대분류' in tdf.columns:
                tdf = tdf[tdf['대분류'] == sel_c1]
            if sel_c2 != '전체' and filter2 is not None:
                tdf = tdf[tdf[filter2] == sel_c2]
            if only_qty:
                tdf = tdf[tdf['재고'] > 0]

            # ── 소계 KPI ─────────────────────────────────────────────
            t1, t2, t3, t4 = st.columns(4)
            t1.metric("항목 수",   f"{len(tdf):,} 건")
            t2.metric("신품 합계", f"{int(tdf['신품'].sum()):,} 개")
            t3.metric("구품 합계", f"{int(tdf['구품'].sum()):,} 개")
            t4.metric("재고 합계", f"{int(tdf['재고'].sum()):,} 개")

            # ── 테이블 출력 ───────────────────────────────────────────
            show_cols = [c for c in display_cols if c in tdf.columns]
            disp = tdf[show_cols].copy().reset_index(drop=True)
            for c in ['이월(신품)','이월(구품)','이월(재고)','신품','구품','재고','전체']:
                if c in disp.columns:
                    disp[c] = disp[c].apply(fmt)

            st.dataframe(
                disp,
                use_container_width=True,
                height=480,
                column_config={
                    '자재코드': st.column_config.TextColumn('자재코드'),
                    '신품':     st.column_config.TextColumn('신품',  help='신품 수량'),
                    '구품':     st.column_config.TextColumn('구품',  help='구품(양호) 수량'),
                    '재고':     st.column_config.TextColumn('재고',  help='신품+구품 합계'),
                }
            )

            # 엑셀 다운로드
            dl_df = tdf[show_cols].copy().reset_index(drop=True)
            import io as _io
            buf_xl = _io.BytesIO()
            with pd.ExcelWriter(buf_xl, engine='openpyxl') as writer:
                dl_df.to_excel(writer, index=False, sheet_name='조회결과')
            buf_xl.seek(0)
            from datetime import datetime
            today = datetime.now().strftime('%Y%m%d')
            file_label = kw if kw else tab_cfg['sheet']
            dl_filename = f"{file_label}_{today}.xlsx"
            st.download_button(
                label=f"⬇️ 엑셀 다운로드 ({len(dl_df):,}건)",
                data=buf_xl,
                file_name=dl_filename,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key=f"dl_{tab_key}"
            )
            st.caption(
                f"표시 {len(tdf):,}건  |  전체 {len(merged):,}건  |  "
                f"재고 보유 {int((merged['재고'] > 0).sum()):,}건"
            )