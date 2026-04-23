import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os

# ─── 페이지 기본 설정 ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="롯데백화점몰 | 쿠폰 업로드 Pro",
    page_icon="🏬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── 글로벌 CSS: 롯데백화점몰 디자인 시스템 ─────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@300;400;600;700&family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');

/* ══════════════════════════════════════
   기본 리셋 & 전체 배경
══════════════════════════════════════ */
html, body, [class*="css"] {
    font-family: 'Noto Sans KR', sans-serif !important;
}
.stApp {
    background-color: #fcfbf9;
}
.main .block-container {
    padding: 2.2rem 3rem 4rem 3rem;
    max-width: 1180px;
}

/* ══════════════════════════════════════
   헤더
══════════════════════════════════════ */
.ld-header {
    background: #1a1a2e;
    color: white;
    padding: 32px 40px;
    margin-bottom: 32px;
    display: flex;
    align-items: center;
    gap: 24px;
    position: relative;
    overflow: hidden;
    border-bottom: 3px solid #b8965a;
    border-radius: 6px;
}
.ld-header::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
    background:
        linear-gradient(135deg, rgba(184,150,90,0.08) 0%, transparent 60%),
        repeating-linear-gradient(
            45deg,
            transparent,
            transparent 40px,
            rgba(184,150,90,0.025) 40px,
            rgba(184,150,90,0.025) 41px
        );
    pointer-events: none;
}
.ld-header-left {
    display: flex;
    flex-direction: column;
    flex: 1;
}
.ld-header-eyebrow {
    font-size: 0.7rem;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #b8965a;
    font-weight: 500;
    margin-bottom: 8px;
    font-family: 'Noto Sans KR', sans-serif;
}
.ld-header h1 {
    margin: 0;
    font-size: 1.7rem;
    font-weight: 700;
    letter-spacing: -0.3px;
    color: white !important;
    font-family: 'Noto Serif KR', serif !important;
    line-height: 1.2;
}
.ld-header-sub {
    margin: 8px 0 0 0;
    font-size: 0.82rem;
    color: rgba(255,255,255,0.55);
    font-weight: 300;
    letter-spacing: 0.3px;
}
.ld-header-badge {
    background: transparent;
    border: 1px solid rgba(184,150,90,0.6);
    color: #b8965a;
    padding: 6px 16px;
    border-radius: 4px;
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 2px;
    text-transform: uppercase;
    white-space: nowrap;
}

/* ══════════════════════════════════════
   안내 메시지 커스텀
══════════════════════════════════════ */
div[data-testid="stAlert"] {
    background: #fffdf8 !important;
    border: 1px solid #e8dcc8 !important;
    border-left: 3px solid #b8965a !important;
    border-radius: 4px !important;
    color: #5a4a35 !important;
    font-size: 0.83rem !important;
}

/* ══════════════════════════════════════
   섹션 카드
══════════════════════════════════════ */
.ld-section {
    background: white;
    border-radius: 8px;
    padding: 28px 32px;
    margin-bottom: 16px;
    border: 1px solid #e8e4de;
    box-shadow: 0 4px 12px rgba(0,0,0,0.03);
}
.ld-section-title {
    font-family: 'Noto Serif KR', serif;
    font-size: 0.98rem;
    font-weight: 600;
    color: #1a1a2e;
    margin-bottom: 22px;
    padding-bottom: 14px;
    border-bottom: 1px solid #f0ece6;
    display: flex;
    align-items: center;
    gap: 10px;
    letter-spacing: -0.2px;
}
.ld-section-title::before {
    content: "";
    display: inline-block;
    width: 3px;
    height: 16px;
    background: linear-gradient(180deg, #b8965a, #d4b47a);
    border-radius: 4px;
    flex-shrink: 0;
}

/* ══════════════════════════════════════
   입력 필드
══════════════════════════════════════ */
.stSelectbox label,
.stNumberInput label,
.stTextInput label,
.stDateInput label,
.stTimeInput label,
.stMultiSelect label,
.stRadio label:first-of-type {
    font-size: 0.77rem !important;
    font-weight: 600 !important;
    color: #7a6e62 !important;
    letter-spacing: 0.8px !important;
    text-transform: uppercase !important;
}
.stTextInput > div > div > input,
.stNumberInput > div > div > input {
    border: 1px solid #ddd8d0 !important;
    border-radius: 3px !important;
    font-size: 0.88rem !important;
    background: #fdfcfb !important;
    color: #1a1a2e !important;
    padding: 9px 13px !important;
    transition: border-color 0.2s, box-shadow 0.2s;
}
.stTextInput > div > div > input:focus,
.stNumberInput > div > div > input:focus {
    border-color: #b8965a !important;
    box-shadow: 0 0 0 3px rgba(184,150,90,0.12) !important;
    background: white !important;
}
.stSelectbox > div > div {
    border: 1px solid #ddd8d0 !important;
    border-radius: 3px !important;
    background: #fdfcfb !important;
}
.stSelectbox > div > div:focus-within {
    border-color: #b8965a !important;
    box-shadow: 0 0 0 3px rgba(184,150,90,0.12) !important;
}

/* ══════════════════════════════════════
   메인 CTA 버튼
══════════════════════════════════════ */
.stButton > button {
    width: 100%;
    border-radius: 6px;
    height: 3.4em;
    background: #1a1a2e !important;
    color: white !important;
    font-size: 0.92rem !important;
    font-weight: 600 !important;
    letter-spacing: 1.5px !important;
    text-transform: uppercase !important;
    border: 1px solid #b8965a !important;
    box-shadow: 0 2px 12px rgba(26,26,46,0.18) !important;
    transition: all 0.25s ease !important;
    font-family: 'Noto Sans KR', sans-serif !important;
}
.stButton > button:hover {
    background: #b8965a !important;
    border-color: #b8965a !important;
    box-shadow: 0 4px 20px rgba(184,150,90,0.35) !important;
    transform: translateY(-1px) !important;
}
.stButton > button:active {
    transform: translateY(0) !important;
}

/* ══════════════════════════════════════
   다운로드 버튼
══════════════════════════════════════ */
.stDownloadButton > button {
    width: 100%;
    border-radius: 6px;
    background: white !important;
    color: #1a1a2e !important;
    height: 3em;
    font-size: 0.8rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.5px !important;
    border: 1px solid #c4bdb4 !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06) !important;
    transition: all 0.2s ease !important;
}
.stDownloadButton > button:hover {
    background: #1a1a2e !important;
    color: white !important;
    border-color: #1a1a2e !important;
    box-shadow: 0 3px 12px rgba(26,26,46,0.2) !important;
}

/* ══════════════════════════════════════
   컬럼 간격
══════════════════════════════════════ */
[data-testid="column"] {
    padding-left: 8px !important;
    padding-right: 8px !important;
}

/* ══════════════════════════════════════
   익스팬더
══════════════════════════════════════ */
div[data-testid="stExpander"] {
    border: 1px solid #e8e4de !important;
    border-radius: 4px !important;
    background: #fdfcfb !important;
    box-shadow: none !important;
}
div[data-testid="stExpander"] summary {
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    color: #1a1a2e !important;
    letter-spacing: 0.3px !important;
}

/* ══════════════════════════════════════
   라디오
══════════════════════════════════════ */
.stRadio > div {
    gap: 16px;
}
.stRadio > div label {
    font-size: 0.84rem !important;
    color: #3a3530 !important;
    font-weight: 400 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
}

/* ══════════════════════════════════════
   캡션
══════════════════════════════════════ */
.stCaption {
    font-size: 0.74rem !important;
    color: #aaa49c !important;
}

/* ══════════════════════════════════════
   사이드바
══════════════════════════════════════ */
[data-testid="stSidebar"] {
    background: #1a1a2e !important;
}
[data-testid="stSidebar"] * {
    color: rgba(255,255,255,0.85) !important;
}
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stMarkdown span {
    color: rgba(255,255,255,0.6) !important;
    font-size: 0.8rem !important;
}
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stNumberInput label,
[data-testid="stSidebar"] .stRadio label:first-of-type {
    color: rgba(255,255,255,0.5) !important;
    font-size: 0.72rem !important;
    letter-spacing: 1px !important;
}
[data-testid="stSidebar"] .stMultiSelect > div > div,
[data-testid="stSidebar"] .stNumberInput > div > div > input {
    background: rgba(255,255,255,0.06) !important;
    border: 1px solid rgba(255,255,255,0.15) !important;
    border-radius: 3px !important;
    color: white !important;
}
[data-testid="stSidebar"] hr {
    border-color: rgba(255,255,255,0.1) !important;
}
[data-testid="stSidebar"] .stRadio > div label {
    color: rgba(255,255,255,0.8) !important;
    font-size: 0.84rem !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    font-weight: 400 !important;
}

/* ══════════════════════════════════════
   사이드바 로고
══════════════════════════════════════ */
.sb-logo {
    padding: 20px 0 24px 0;
    border-bottom: 1px solid rgba(184,150,90,0.3);
    margin-bottom: 24px;
}
.sb-logo-main {
    font-family: 'Noto Serif KR', serif;
    font-size: 1.05rem;
    font-weight: 700;
    color: white !important;
    letter-spacing: -0.3px;
    line-height: 1.3;
}
.sb-logo-sub {
    font-size: 0.68rem;
    color: #b8965a !important;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    margin-top: 3px;
}

/* ══════════════════════════════════════
   사이드바 섹션 구분
══════════════════════════════════════ */
.sb-section-label {
    font-size: 0.67rem !important;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #b8965a !important;
    font-weight: 600;
    margin-bottom: 10px;
    padding-bottom: 6px;
    border-bottom: 1px solid rgba(184,150,90,0.25);
}

/* ══════════════════════════════════════
   결과 통계 박스
══════════════════════════════════════ */
.stat-card {
    background: white;
    border: 1px solid #e8e4de;
    border-top: 2px solid #b8965a;
    border-radius: 6px;
    padding: 20px 18px;
    text-align: center;
    box-shadow: 0 1px 6px rgba(0,0,0,0.05);
}
.stat-card .num {
    font-family: 'Noto Serif KR', serif;
    font-size: 2rem;
    font-weight: 700;
    color: #1a1a2e;
    line-height: 1;
    letter-spacing: -1px;
}
.stat-card .unit {
    font-size: 0.78rem;
    color: #b8965a;
    font-weight: 600;
    margin-left: 2px;
}
.stat-card .label {
    font-size: 0.73rem;
    color: #aaa49c;
    margin-top: 6px;
    font-weight: 400;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}

/* ══════════════════════════════════════
   다운로드 섹션 헤더
══════════════════════════════════════ */
.dl-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin: 24px 0 14px 0;
    padding-bottom: 12px;
    border-bottom: 1px solid #e8e4de;
}
.dl-header-title {
    font-family: 'Noto Serif KR', serif;
    font-size: 0.95rem;
    font-weight: 600;
    color: #1a1a2e;
    display: flex;
    align-items: center;
    gap: 8px;
}
.dl-header-badge {
    background: #1a1a2e;
    color: #b8965a;
    padding: 3px 12px;
    border-radius: 4px;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 1px;
}

/* ══════════════════════════════════════
   푸터
══════════════════════════════════════ */
.ld-footer {
    text-align: center;
    margin-top: 52px;
    padding-top: 20px;
    border-top: 1px solid #e8e4de;
    color: #c4bdb4;
    font-size: 0.72rem;
    letter-spacing: 1.5px;
    text-transform: uppercase;
}
.ld-footer span {
    color: #b8965a;
}

[data-testid="stSidebar"] .stNumberInput input {
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
}
/* Ensure +/- buttons in sidebar are visible */
[data-testid="stSidebar"] button[title="Step up"],
[data-testid="stSidebar"] button[title="Step down"] {
    color: #ffffff !important;
    background: transparent !important;
}
[data-testid="stSidebar"] button[title="Step up"] svg,
[data-testid="stSidebar"] button[title="Step down"] svg {
    fill: #ffffff !important;
}



/* 강력한 사이드바 인풋(텍스트, 숫자) 색상 강제 지정 및 넓이 확보 */
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] input[type="number"],
[data-testid="stSidebar"] input[type="text"] {
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    caret-color: #ffffff !important;
    background-color: rgba(255, 255, 255, 0.1) !important;
    border: 1px solid rgba(255, 255, 255, 0.3) !important;
    border-radius: 4px !important;
    padding: 2px 4px !important;
    font-size: 0.9rem !important;
    letter-spacing: -0.5px !important;
}

[data-testid="stSidebar"] div[data-baseweb="input"],
[data-testid="stSidebar"] div[data-baseweb="base-input"] {
    background-color: transparent !important;
    padding: 0 !important;
}

/* X (Clear) 버튼 숨기기 */
[data-testid="stSidebar"] button[aria-label="Clear input"],
[data-testid="stSidebar"] button[aria-label="Clear value"],
[data-testid="stSidebar"] div[data-baseweb="input"] > div > button {
    display: none !important;
}

/* 숫자 조절(+, -) 버튼 상시 노출 최적화 (Streamlit 내부 클래스 직접 타겟팅) */
[data-testid="stSidebar"] div[data-testid="InputInstructions"] {
    display: none !important; /* 내부 안내 문구 숨김 */
}

/* step-up, step-down 버튼을 항상 강제 표시 */
[data-testid="stSidebar"] button[data-testid="step-up"],
[data-testid="stSidebar"] button[data-testid="step-down"],
[data-testid="stSidebar"] button.step-up,
[data-testid="stSidebar"] button.step-down {
    background-color: rgba(255, 255, 255, 0.2) !important;
    color: #ffffff !important;
    opacity: 1 !important; 
    visibility: visible !important;
    display: flex !important;
    transform: none !important;
}

[data-testid="stSidebar"] button[data-testid="step-up"] svg,
[data-testid="stSidebar"] button[data-testid="step-down"] svg {
    fill: #ffffff !important;
    color: #ffffff !important;
    opacity: 1 !important;
    visibility: visible !important;
}
</style>
""", unsafe_allow_html=True)

# ─── 헤더 ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="ld-header">
    <div class="ld-header-left">
        <div class="ld-header-eyebrow">Lotte Department Store Mall</div>
        <h1>쿠폰 대량 업로드</h1>
        <p class="ld-header-sub">상품 조건을 설정하고 쿠폰 업로드용 엑셀 파일을 자동으로 생성합니다</p>
    </div>
    <div class="ld-header-badge">Pro Tool</div>
</div>
""", unsafe_allow_html=True)

st.info("좌측 필터에서 조건을 설정한 뒤, 하단의 엑셀 파일 추출 버튼을 눌러주세요.")

# ─── 데이터 로드 ──────────────────────────────────────────────────────────────
DEFAULT_FILE = "product_dummy_data_30k.csv"

@st.cache_data
def load_data(file_path_or_buffer):
    try:
        if isinstance(file_path_or_buffer, str):
            return pd.read_csv(file_path_or_buffer, encoding='utf-8-sig')
        else:
            return pd.read_csv(file_path_or_buffer, encoding='utf-8-sig') \
                if file_path_or_buffer.name.endswith('csv') \
                else pd.read_excel(file_path_or_buffer)
    except Exception:
        return None

df_raw = load_data(DEFAULT_FILE) if os.path.exists(DEFAULT_FILE) else None

# ─── 사이드바 ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sb-logo">
        <div class="sb-logo-main">🏬 롯데백화점몰</div>
        <div class="sb-logo-sub">Premium Coupon Upload</div>
    </div>
    """, unsafe_allow_html=True)

    if df_raw is not None:
        st.markdown('<p class="sb-section-label">상품 필터</p>', unsafe_allow_html=True)
        sel_store    = st.multiselect("점포 (상위거래처)",  sorted(df_raw['상위거래처'].unique()))
        sel_md_group = st.multiselect("상위 MD 상품군",    sorted(df_raw['상위MD상품군명'].unique()))
        sel_md       = st.multiselect("백화점 MD",         sorted(df_raw['백화점MD'].unique()))
        sel_brand    = st.multiselect("브랜드명",           sorted(df_raw['브랜드명'].unique()))

        st.markdown("---")
        st.markdown('<p class="sb-section-label">마진율 필터</p>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1: sel_min_margin = st.number_input("이상", min_value=0, max_value=100, value=None, step=1)
        with c2: sel_max_margin = st.number_input("이하", min_value=0, max_value=100, value=None, step=1)

        st.markdown("---")
        st.markdown('<p class="sb-section-label">상품 상태</p>', unsafe_allow_html=True)
        status_option = st.radio("", ("전체", "전시", "미전시"), horizontal=True)
    else:
        st.warning("`product_dummy_data_30k.csv` 파일을 찾을 수 없습니다.")
        sel_store = sel_md_group = sel_md = sel_brand = []
        sel_min_margin, sel_max_margin = 10, 40
        status_option = "전체"

# ─── 섹션 1: 기본 발행 정책 ──────────────────────────────────────────────────
st.markdown('<div class="ld-section"><div class="ld-section-title">기본 발행 정책</div>', unsafe_allow_html=True)
col1, col2, col3 = st.columns(3)
with col1:
    store_range    = st.selectbox("매장 범위", ["A (전채널)", "M (본매장)", "O (외부매장)"])
with col2:
    discount_type  = st.selectbox("할인 유형",  ["10 (정률)", "30 (정액)"])
with col3:
    discount_value = st.number_input("할인액 / 율", min_value=0, value=10)
st.markdown('</div>', unsafe_allow_html=True)

# ─── 섹션 2: 비용 분담 & 요일 ────────────────────────────────────────────────
st.markdown('<div class="ld-section"><div class="ld-section-title">비용 분담 및 요일 설정</div>', unsafe_allow_html=True)
col1, col2, col3 = st.columns(3)
with col1:
    vendor_share  = st.number_input("거래처 분담율 (%)", 0, 100, 0)
with col2:
    partner_share = st.number_input("제휴사 분담율 (%)", 0, 100, 0)
with col3:
    day_setting_input = st.text_input("요일 설정 (월~일, O/X)", value="OOOOOOO")
    st.caption("예시 — 금·토·일만 적용: XXXXOOO")
    day_setting  = day_setting_input.upper()
    is_day_valid = len(day_setting) == 7 and all(c in "OX" for c in day_setting)
    if not is_day_valid:
        st.error("'O'와 'X'만 사용하여 정확히 7자리를 입력해주세요.")
st.markdown('</div>', unsafe_allow_html=True)

# ─── 섹션 3: 행사 기간 ───────────────────────────────────────────────────────
st.markdown('<div class="ld-section"><div class="ld-section-title">행사 기간 설정</div>', unsafe_allow_html=True)
col_date, col_time = st.columns(2)
with col_date:
    with st.expander("날짜 설정", expanded=True):
        d1, d2 = st.columns(2)
        start_dt = d1.date_input("시작일", datetime.now())
        end_dt   = d2.date_input("종료일", datetime.now())
with col_time:
    with st.expander("시간 설정", expanded=True):
        t1, t2 = st.columns(2)
        start_tm = t1.time_input("시작 시간", datetime.strptime("00:00", "%H:%M").time())
        end_tm   = t2.time_input("종료 시간", datetime.strptime("23:59", "%H:%M").time())
st.markdown('</div>', unsafe_allow_html=True)

# ─── CTA 버튼 ─────────────────────────────────────────────────────────────────
st.markdown("<div style='margin: 28px 0 8px 0;'></div>", unsafe_allow_html=True)
extract_btn = st.button("엑셀 파일 추출 및 다운로드 메뉴 생성")

# ─── 실행 로직 ────────────────────────────────────────────────────────────────
if extract_btn:
    if not is_day_valid:
        st.warning("요일 설정을 올바르게 입력한 후 다시 시도해주세요.")
    elif df_raw is None:
        st.error("데이터 파일을 찾을 수 없습니다. CSV 파일을 확인해주세요.")
    else:
        df_f = df_raw.copy()
        if sel_store:    df_f = df_f[df_f['상위거래처'].isin(sel_store)]
        if sel_md_group: df_f = df_f[df_f['상위MD상품군명'].isin(sel_md_group)]
        if sel_md:       df_f = df_f[df_f['백화점MD'].isin(sel_md)]
        if sel_brand:    df_f = df_f[df_f['브랜드명'].isin(sel_brand)]
        if status_option != "전체": df_f = df_f[df_f['상품상태'] == status_option]
                
        if sel_min_margin is not None:
            df_f = df_f[df_f['마진율'] >= sel_min_margin]
        if sel_max_margin is not None:
            df_f = df_f[df_f['마진율'] <= sel_max_margin]


        total_count = len(df_f)
        CHUNK_SIZE  = 5000
        num_chunks  = max(1, (total_count + CHUNK_SIZE - 1) // CHUNK_SIZE)

        # ── 통계 카드 ──
        st.markdown("<div style='margin-top:24px;'></div>", unsafe_allow_html=True)
        s1, s2, s3 = st.columns(3)
        with s1:
            st.markdown(f"""
            <div class="stat-card">
                <div class="num">{total_count:,}<span class="unit">건</span></div>
                <div class="label">추출된 상품 수</div>
            </div>""", unsafe_allow_html=True)
        with s2:
            st.markdown(f"""
            <div class="stat-card">
                <div class="num">{num_chunks}<span class="unit">개</span></div>
                <div class="label">분할 파일 수</div>
            </div>""", unsafe_allow_html=True)
        with s3:
            st.markdown(f"""
            <div class="stat-card">
                <div class="num">{CHUNK_SIZE:,}<span class="unit">건</span></div>
                <div class="label">파일당 최대 건수</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)

        if total_count > 0:
            st.success(f"총 **{total_count:,}**개 상품이 추출되었습니다.")

            start_str = datetime.combine(start_dt, start_tm).strftime('%Y%m%d%H%M')
            end_str   = datetime.combine(end_dt,   end_tm).strftime('%Y%m%d%H%M')

            df_upload = pd.DataFrame({
                '상품번호':         df_f['상품번호'],
                '매장범위':         store_range[0],
                '행사시작일':       start_str,
                '행사종료일':       end_str,
                '할인유형':         discount_type.split(' ')[0],
                '할인액':           discount_value,
                '거래처분담율':     vendor_share,
                '제휴사분담율':     partner_share,
                '사용요일':         day_setting,
                '시작시간':         '0000',
                '종료시간':         '2359',
                '요일/시간 할인율': "",
            })

            st.markdown(f"""
            <div class="dl-header">
                <div class="dl-header-title">분할 파일 다운로드</div>
                <div class="dl-header-badge">{num_chunks} FILES</div>
            </div>
            """, unsafe_allow_html=True)

            cols = st.columns(4)
            for idx in range(num_chunks):
                start_idx = idx * CHUNK_SIZE
                end_idx   = min(start_idx + CHUNK_SIZE, total_count)
                if start_idx >= total_count:
                    break

                chunk_df = df_upload.iloc[start_idx:end_idx]
                output   = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    chunk_df.to_excel(writer, index=False)

                with cols[idx % 4]:
                    st.download_button(
                        label=f"Part {idx+1}  ·  {start_idx+1:,}–{end_idx:,}",
                        data=output.getvalue(),
                        file_name=f"coupon_part{idx+1}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"btn_{idx}",
                    )
        else:
            st.warning("선택한 조건에 해당하는 상품이 없습니다. 필터 조건을 조정해보세요.")

# ─── 푸터 ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="ld-footer">
    <span>Lotte Department Store Mall</span> &nbsp;·&nbsp; Premium Coupon Upload &nbsp;·&nbsp; Internal Use Only
</div>
""", unsafe_allow_html=True)
