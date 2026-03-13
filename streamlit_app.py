import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

st.set_page_config(
    page_title="자재 현황 대시보드",
    page_icon="📦",
    layout="wide"
)

# ── 스타일 (Warm White 테마) ────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #FAFAF8; }
[data-testid="stSidebar"]          { background: #F5F3EE; border-right: 1px solid #E8E4DC; }
.kpi-card {
    background: #FFFFFF;
    border: 0.5px solid #E8E4DC; border-radius: 12px;
    padding: 20px 24px; text-align: center;
}
.kpi-label { color: #7C7268; font-size: 13px; margin-bottom: 6px; }
.kpi-value { color: #1A1A1A; font-size: 32px; font-weight: 600; }
.kpi-unit  { color: #A89E94; font-size: 13px; margin-top: 4px; }
.section-title {
    color: #3D3530; font-size: 15px; font-weight: 600;
    padding: 8px 0 5px 0;
    border-bottom: 2px solid #B45309; margin-bottom: 12px;
    display: inline-block;
}
.status-ok   { color: #166534; font-weight: 600; }
.status-warn { color: #B45309; font-weight: 600; }
[data-testid="stTabs"] button { font-size: 14px !important; font-weight: 600 !important; color: #3D3530 !important; }
[data-testid="stTabs"] button[aria-selected="true"] { border-bottom-color: #B45309 !important; color: #1A1A1A !important; }
</style>
""", unsafe_allow_html=True)

# ── 탭 설정 ─────────────────────────────────────────────────────
# 컬럼: (내부 DataFrame 컬럼명, 화면 표시 헤더)
# 필터2차: 2차 드롭다운 기준 컬럼
TAB_CONFIG = [
    {
        "label":      "📡 5G 물자",
        "대분류목록": ["5G"],
        "컬럼":       [("대분류","대분류"),("중분류","중분류"),("소분류","소분류"),
                       ("자재코드","자재코드"),("품명","품명"),("제조사","제조사"),
                       ("신품","신품"),("구품","구품"),("재고","재고")],
        "필터2차":    "중분류",
    },
    {
        "label":      "📶 5G·LTE 물자",
        "대분류목록": ["5G", "LTE"],
        # 중분류 없이 소분류만 표시
        "컬럼":       [("대분류","대분류"),("소분류","소분류"),
                       ("자재코드","자재코드"),("품명","품명"),("제조사","제조사"),
                       ("신품","신품"),("구품","구품"),("재고","재고")],
        "필터2차":    "소분류",
    },
    {
        "label":      "🔧 RRU·MiBos·W·설비물자",
        "대분류목록": ["RRU", "MiBOS", "MiBos", "W", "설비물자", " 설비물자",
                       "LTE 정류기", "LTE 축전지", "SF-W15", "SF-W20",
                       "납축전지", "대형리튬축전지", "대형정류기", "함체", "외함체",
                       "RACK", "리튬축전지", "5G정류기", "DUH RACK", "표준함체",
                       "RMS", "기지국 장비"],
        # 소분류/제조사 명칭 통일
        "컬럼":       [("대분류","대분류"),("중분류","중분류"),("소분류","소분류"),
                       ("자재코드","자재코드"),("품명","품명"),("제조사","제조사"),
                       ("신품","신품"),("구품","구품"),("재고","재고")],
        "필터2차":    "중분류",
    },
]

# ── 매핑 파일 경로 ──────────────────────────────────────────────
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
MAPPING_PATH = os.path.join(BASE_DIR, "data", "material_mapping.csv")

# ── 헬퍼 ───────────────────────────────────────────────────────
def fmt(val):
    return '-' if val == 0 else f'{val:,}'

# ── 캐시 함수 ───────────────────────────────────────────────────
@st.cache_data
def load_mapping(path):
    expected = ['대분류','중분류','소분류','자재코드','품명','제조사']
    try:
        mp = pd.read_csv(path, encoding='utf-8-sig', dtype=str)
        if len(mp.columns) < 6:
            raise ValueError(f"컬럼 수 부족: {len(mp.columns)}개")
        mp = mp.iloc[:, :6]
        mp.columns = expected
    except Exception as e:
        st.warning(f"매핑 파일을 읽을 수 없습니다: {e}")
        return pd.DataFrame(columns=expected)
    mp = mp.dropna(subset=['자재코드'])
    mp['자재코드'] = pd.to_numeric(mp['자재코드'], errors='coerce').astype('Int64')
    mp['대분류']   = mp['대분류'].astype(str).str.strip()
    return mp.dropna(subset=['자재코드'])

@st.cache_data
def parse_excel(file_bytes, file_name):
    import io
    buf = io.BytesIO(file_bytes)
    raw = pd.read_excel(buf, header=None)
    # 헤더 행 자동 탐색: '순번' 또는 '자재코드'가 있는 행
    header_row = 0
    for i in range(min(5, len(raw))):
        row_vals = raw.iloc[i].astype(str).tolist()
        if any('순번' in v or '자재코드' in v for v in row_vals):
            header_row = i
            break
    buf.seek(0)
    df = pd.read_excel(buf, header=header_row)
    if len(df.columns) < 12:
        raise ValueError(f"컬럼 수가 {len(df.columns)}개입니다. 12개 이상이어야 합니다 (헤더 행 구조 확인)")
    df = df.iloc[:, :12]
    df.columns = ['순번','사업년도','지역본부','군','업체명','자재분류',
                  '자재코드','자재명','FULL자재명','신품','구품_양호','구품_불량']
    df['신품']      = pd.to_numeric(df['신품'],      errors='coerce').fillna(0).astype(int)
    df['구품_양호'] = pd.to_numeric(df['구품_양호'], errors='coerce').fillna(0).astype(int)
    df['구품']      = df['구품_양호']   # K열(양호)만 사용
    df['재고']      = df['신품'] + df['구품']
    df['자재코드']  = pd.to_numeric(df['자재코드'], errors='coerce').astype('Int64')
    df['파일명']    = file_name
    return df.dropna(subset=['순번'])

# ── 사이드바 ────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📁 파일 관리")

    st.markdown("#### 자재코드 매핑")
    if os.path.exists(MAPPING_PATH):
        mp_all = load_mapping(MAPPING_PATH)
        st.markdown('<span class="status-ok">✅ 매핑 파일 로드됨</span>', unsafe_allow_html=True)
        st.caption(f"{len(mp_all):,}개 자재코드 | 대분류: {', '.join(sorted(mp_all['대분류'].unique()))}")
    else:
        mp_all = None
        st.markdown('<span class="status-warn">⚠️ data/material_mapping.csv 없음</span>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### 수불부 현황 업로드")
    st.caption("여러 파일 동시 업로드 가능")
    uploaded_files = st.file_uploader(
        "엑셀 파일 선택",
        type=["xlsx","xls"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )


# ── 메인 헤더 ───────────────────────────────────────────────────
st.markdown("# 📦 자재 현황 대시보드")

if not uploaded_files:
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        st.info("**① 수불부 현황 엑셀을 업로드하세요**\n\n좌측 사이드바에서 파일 선택 시 대시보드가 생성됩니다.")
    with c2:
        if mp_all is not None:
            st.success(f"**② 자재코드 매핑 준비 완료**\n\n{len(mp_all):,}개 자재코드 등록됨")
        else:
            st.warning("**② 코드_매핑.xlsx를 앱 폴더에 넣어주세요**")
    st.stop()

# ── 수불부 파싱 ─────────────────────────────────────────────────
all_dfs = []
parse_errors = []
for f in uploaded_files:
    try:
        parsed = parse_excel(f.read(), f.name)
        if parsed is not None and len(parsed) > 0:
            all_dfs.append(parsed)
        else:
            parse_errors.append(f"{f.name}: 데이터가 없습니다 (헤더 행 구조를 확인하세요)")
    except Exception as e:
        parse_errors.append(f"{f.name}: {e}")

if parse_errors:
    for msg in parse_errors:
        st.error(f"❌ {msg}")

if not all_dfs:
    st.warning("업로드된 파일을 읽을 수 없습니다. 엑셀 파일 형식(헤더 2행 구조)을 확인해주세요.")
    st.stop()

df_raw = pd.concat(all_dfs, ignore_index=True)

# ── 전체 필터 없음 (탭별 필터로 대체) ─────────────────────────
fdf = df_raw.copy()

file_tag = " | ".join(f"📄 {f}" for f in df_raw['파일명'].unique())
st.markdown(f"<small style='color:#A89E94'>{file_tag}</small>", unsafe_allow_html=True)
st.markdown("---")

# ── KPI ────────────────────────────────────────────────────────
total_rows  = len(fdf)
total_new   = int(fdf['신품'].sum())
total_used  = int(fdf['구품'].sum())
total_stock = int(fdf['재고'].sum())
has_qty_cnt = int((fdf['재고'] > 0).sum())

def kpi(col, label, value, unit=""):
    col.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value:,}</div>
        <div class="kpi-unit">{unit}</div>
    </div>""", unsafe_allow_html=True)

k1,k2,k3,k4,k5 = st.columns(5)
kpi(k1, "전체 자재 항목", total_rows,    "건")
kpi(k2, "신품 수량",       total_new,     "개")
kpi(k3, "구품 수량",       total_used,    "개")
kpi(k4, "총 재고",         total_stock,   "개")
kpi(k5, "재고 보유 항목",  has_qty_cnt,   "건")
st.markdown("<br>", unsafe_allow_html=True)

# ── 공통 차트 레이아웃 ──────────────────────────────────────────
LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
    font_color='#3D3530', margin=dict(t=20, b=20)
)

# ── 차트 Row 1 ──────────────────────────────────────────────────
c1, c2 = st.columns(2)

with c1:
    st.markdown('<div class="section-title">📊 신품 / 구품 비율</div>', unsafe_allow_html=True)
    if total_stock > 0:
        pie_df = pd.DataFrame({'구분':['신품','구품'], '수량':[total_new, total_used]})
        fig = px.pie(pie_df, names='구분', values='수량',
                     color_discrete_sequence=['#B45309','#92400E'], hole=0.45)
        fig.update_traces(textinfo='label+percent', textfont_size=13)
        fig.update_layout(**LAYOUT, legend=dict(font=dict(color='#c5cae9')))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("수량 데이터가 없습니다.")

with c2:
    st.markdown('<div class="section-title">🏢 업체별 재고 현황 (TOP 10)</div>', unsafe_allow_html=True)
    biz = (fdf[fdf['재고']>0].groupby('업체명')[['신품','구품']]
           .sum().reset_index())
    biz['재고'] = biz['신품'] + biz['구품']
    biz = biz.sort_values('재고', ascending=False).head(10)
    if not biz.empty:
        fig = go.Figure()
        fig.add_bar(x=biz['업체명'], y=biz['신품'], name='신품', marker_color='#B45309')
        fig.add_bar(x=biz['업체명'], y=biz['구품'], name='구품', marker_color='#D97706')
        fig.update_layout(barmode='stack', **LAYOUT,
                          xaxis=dict(gridcolor='#E8E4DC'), yaxis=dict(gridcolor='#E8E4DC'),
                          legend=dict(font=dict(color='#c5cae9')))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("업체 데이터가 없습니다.")

# ── 차트 Row 2 ──────────────────────────────────────────────────
c3, c4 = st.columns(2)

with c3:
    st.markdown('<div class="section-title">📦 재고 TOP 15 자재</div>', unsafe_allow_html=True)
    top = (fdf[fdf['재고']>0].groupby('자재명')['재고'].sum()
           .reset_index().sort_values('재고', ascending=True).tail(15))
    if not top.empty:
        top['자재명_short'] = top['자재명'].str[:25]
        fig = px.bar(top, x='재고', y='자재명_short', orientation='h',
                     color='재고', color_continuous_scale='YlOrBr')
        fig.update_layout(**LAYOUT, xaxis=dict(gridcolor='#E8E4DC'),
                          yaxis=dict(gridcolor='#E8E4DC'), coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("재고 데이터가 없습니다.")

with c4:
    st.markdown('<div class="section-title">🗂️ 자재분류별 항목 수</div>', unsafe_allow_html=True)
    mc = fdf.groupby('자재분류').size().reset_index(name='항목수')
    fig = px.bar(mc, x='자재분류', y='항목수', color='항목수',
                 color_continuous_scale='YlOrBr', text='항목수')
    fig.update_traces(textposition='outside')
    fig.update_layout(**LAYOUT, margin=dict(t=30,b=20),
                      xaxis=dict(gridcolor='#E8E4DC'), yaxis=dict(gridcolor='#E8E4DC'),
                      coloraxis_showscale=False)
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")

# ── 수불부 자재코드별 수량 합산 (조인용) ────────────────────────
qty_df = (fdf.groupby('자재코드')[['신품','구품','재고']]
          .sum().reset_index())

# ── 상세 테이블 (탭 3개) ────────────────────────────────────────
st.markdown('<div class="section-title">📋 자재 상세 목록</div>', unsafe_allow_html=True)

tab_objs = st.tabs([t["label"] for t in TAB_CONFIG])

for tab_obj, tab_cfg in zip(tab_objs, TAB_CONFIG):
    with tab_obj:
        if mp_all is None:
            st.warning("data/material_mapping.csv가 없습니다. data 폴더에 파일을 넣어주세요.")
            continue

        # 해당 탭 대분류만 필터 (대분류는 strip 처리된 상태)
        mp_tab = mp_all[mp_all['대분류'].isin(tab_cfg["대분류목록"])].copy()

        if mp_tab.empty:
            st.info(
                f"매핑 파일에 해당 대분류가 없습니다: {', '.join(tab_cfg['대분류목록'])}\n\n"
                "코드_매핑.xlsx에 해당 대분류 행을 추가하면 자동으로 표시됩니다."
            )
            continue

        # 자재코드 기준 LEFT JOIN
        merged = pd.merge(mp_tab, qty_df, on='자재코드', how='left')
        merged['신품'] = merged['신품'].fillna(0).astype(int)
        merged['구품'] = merged['구품'].fillna(0).astype(int)
        merged['재고'] = merged['재고'].fillna(0).astype(int)

        # ── 탭 내 필터 UI (메인 화면) ────────────────────────
        tab_key  = tab_cfg["label"]
        필터2차  = tab_cfg["필터2차"]

        st.markdown(
            "<div style='background:#FEF3C7;border:0.5px solid #F59E0B;border-radius:10px;"            "padding:14px 18px;margin-bottom:14px'>",
            unsafe_allow_html=True
        )
        f1, f2, f3, f4 = st.columns([3, 2, 2, 1])
        with f1:
            kw = st.text_input("🔍 품명 검색", placeholder="품명 키워드를 입력하세요",
                               key=f"kw_{tab_key}")
        with f2:
            c1_opts = ['전체'] + sorted(merged['대분류'].dropna().unique().tolist())
            sel_c1  = st.selectbox("대분류", c1_opts, key=f"c1_{tab_key}")
        with f3:
            src2    = merged if sel_c1 == '전체' else merged[merged['대분류'] == sel_c1]
            c2_opts = ['전체'] + sorted(src2[필터2차].dropna().unique().tolist())
            sel_c2  = st.selectbox("중분류", c2_opts, key=f"c2_{tab_key}")
        with f4:
            st.markdown("<div style='padding-top:28px'></div>", unsafe_allow_html=True)
            only_qty = st.checkbox("재고만", value=False, key=f"qty_{tab_key}")
        st.markdown("</div>", unsafe_allow_html=True)

        # ── 필터 적용 ─────────────────────────────────────────
        tdf = merged.copy()
        if kw:
            tdf = tdf[tdf['품명'].str.contains(kw, na=False, case=False)]
        if sel_c1 != '전체':
            tdf = tdf[tdf['대분류'] == sel_c1]
        if sel_c2 != '전체':
            tdf = tdf[tdf[필터2차] == sel_c2]
        if only_qty:
            tdf = tdf[tdf['재고'] > 0]

        # ── 탭 내 소계 KPI ────────────────────────────────────
        t1, t2, t3, t4 = st.columns(4)
        t1.metric("항목 수",   f"{len(tdf):,} 건")
        t2.metric("신품 합계", f"{int(tdf['신품'].sum()):,} 개")
        t3.metric("구품 합계", f"{int(tdf['구품'].sum()):,} 개")
        t4.metric("재고 합계", f"{int(tdf['재고'].sum()):,} 개")

        # ── 표시용 DataFrame (탭별 컬럼 구성 적용) ───────────
        내부컬럼 = [c for c, _ in tab_cfg["컬럼"]]
        헤더명   = [h for _, h in tab_cfg["컬럼"]]

        display_df = tdf[내부컬럼].copy().reset_index(drop=True)
        display_df.columns = 헤더명

        # 수량 포맷 (0 → '-')
        for col in ['신품','구품','재고']:
            if col in display_df.columns:
                display_df[col] = display_df[col].apply(fmt)

        st.dataframe(
            display_df,
            use_container_width=True,
            height=480,
            column_config={
                '자재코드': st.column_config.TextColumn('자재코드'),
                '신품':     st.column_config.TextColumn('신품',  help='신품 수량'),
                '구품':     st.column_config.TextColumn('구품',  help='구품(양호) 수량'),
                '재고':     st.column_config.TextColumn('재고',  help='신품+구품 합계'),
            }
        )
        st.caption(
            f"표시 {len(tdf):,}건  |  "
            f"전체 {len(merged):,}건  |  "
            f"재고 보유 {int((merged['재고'] > 0).sum()):,}건"
        )