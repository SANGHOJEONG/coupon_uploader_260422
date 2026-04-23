import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os

# 페이지 기본 설정
st.set_page_config(page_title="롯데온 Pro 쿠폰 업로드 지원(Final)", layout="wide")

# CSS를 활용한 디자인 요소 (버튼 크기 유지 및 간격 축소 최적화)
st.markdown('''
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #ff4b4b; color: white; }
    
    /* 다운로드 버튼: Ver2 수준의 크기를 유지하면서 간격만 좁힘 */
    .stDownloadButton>button { 
        width: 100%; 
        border-radius: 3px; 
        background-color: #007bff; 
        color: white;
        height: 3em; /* 버튼 높이 고정 */
    }
    
    /* 컬럼 간의 불필요한 좌우 여백(Padding)을 최소화하여 간격 축소 */
    [data-testid="column"] {
        padding-left: 5px !important;
        padding-right: 5px !important;
    }
    
    div[data-testid="stExpander"] { border: 1px solid #e0e0e0; border-radius: 10px; background-color: white; }
    </style>
    ''', unsafe_allow_html=True)

st.title("🎫 롯데온 Pro 쿠폰 대량 업로드 지원 App")
st.info("조건을 선택한 후 하단의 '엑셀 파일 추출' 버튼을 눌러주세요.")

# -----------------------------------------------------------------------------
# 1. 데이터 로드 (로컬 우선)
# -----------------------------------------------------------------------------
DEFAULT_FILE = "product_dummy_data_30k.csv"

@st.cache_data
def load_data(file_path_or_buffer):
    try:
        if isinstance(file_path_or_buffer, str):
            return pd.read_csv(file_path_or_buffer, encoding='utf-8-sig')
        else:
            return pd.read_csv(file_path_or_buffer, encoding='utf-8-sig') if file_path_or_buffer.name.endswith('csv') else pd.read_excel(file_path_or_buffer)
    except Exception: return None

df_raw = load_data(DEFAULT_FILE) if os.path.exists(DEFAULT_FILE) else None

# -----------------------------------------------------------------------------
# 2. Sidebar: 필터 설정
# -----------------------------------------------------------------------------
st.sidebar.header("🔍 상품 필터링 조건")

if df_raw is not None:
    sel_store = st.sidebar.multiselect("점포(상위거래처)", sorted(df_raw['상위거래처'].unique()))
    sel_md_group = st.sidebar.multiselect("상위MD상품군", sorted(df_raw['상위MD상품군명'].unique()))
    sel_md = st.sidebar.multiselect("백화점MD", sorted(df_raw['백화점MD'].unique()))
    sel_brand = st.sidebar.multiselect("브랜드명", sorted(df_raw['브랜드명'].unique()))
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("**수익성 필터**")
    c1, c2 = st.sidebar.columns(2)
    with c1: sel_min_margin = st.number_input("마진율 이상", 0, 100, 10)
    with c2: sel_max_margin = st.number_input("마진율 이하", 0, 100, 40)

    status_option = st.sidebar.radio("상품 상태", ("전체", "전시", "미전시"), horizontal=True)
else:
    st.sidebar.warning("`product_dummy_data_30k.csv` 파일을 찾을 수 없습니다.")

# -----------------------------------------------------------------------------
# 3. Main Dashboard
# -----------------------------------------------------------------------------
with st.container():
    st.subheader("⚙️ 쿠폰 발행 정책 및 기간 설정")
    
    st.markdown("**📍 기본 정책**")
    r1_c1, r1_c2, r1_c3 = st.columns(3)
    with r1_c1:
        store_range = st.selectbox("매장범위", ["A (전채널)", "M (본매장)", "O (외부매장)"])
    with r1_c2:
        discount_type = st.selectbox("할인유형", ["10 (정률)", "30 (정액)"])
    with r1_c3:
        discount_value = st.number_input("할인액/율", min_value=0, value=10)
    
    st.markdown("**💰 비용 분담 및 기타 설정**")
    r2_c1, r2_c2, r2_c3 = st.columns(3)
    with r2_c1:
        vendor_share = st.number_input("거래처 분담율 (%)", 0, 100, 0)
    with r2_c2:
        partner_share = st.number_input("제휴사 분담율 (%)", 0, 100, 0)
    with r2_c3:
        day_setting_input = st.text_input("요일 설정", value="OOOOOOO")
        st.caption("💡 금토일만 쿠폰을 붙이고 싶다면 XXXXOOO으로 설정")
        
        # 유효성 검사
        day_setting = day_setting_input.upper()
        is_day_valid = len(day_setting) == 7 and all(c in "OX" for c in day_setting)
        if not is_day_valid:
            st.error("⚠️ 'O'와 'X'만 사용하여 7자리를 입력해주세요.")

    st.markdown("---")
    row2_col1, row2_col2 = st.columns(2)
    
    with row2_col1:
        with st.expander("📅 행사 날짜 설정", expanded=True):
            d_c1, d_c2 = st.columns(2)
            start_dt = d_c1.date_input("시작일", datetime.now())
            end_dt = d_c2.date_input("종료일", datetime.now())
            
    with row2_col2:
        with st.expander("⏰ 행사 시간 설정", expanded=True):
            t_c1, t_c2 = st.columns(2)
            start_tm = t_c1.time_input("시작 시간", datetime.strptime("00:00", "%H:%M").time())
            end_tm = t_c2.time_input("종료 시간", datetime.strptime("23:59", "%H:%M").time())

# -----------------------------------------------------------------------------
# 4. 실행 로직
# -----------------------------------------------------------------------------
st.markdown("###")
extract_btn = st.button("🔍 엑셀 파일 추출 및 다운로드 메뉴 생성")

if extract_btn:
    if not is_day_valid:
        st.warning("요일 설정을 올바르게 입력해야 추출이 가능합니다.")
    elif df_raw is None:
        st.error("데이터 파일을 찾을 수 없습니다.")
    else:
        # 필터링
        df_f = df_raw.copy()
        if sel_store: df_f = df_f[df_f['상위거래처'].isin(sel_store)]
        if sel_md_group: df_f = df_f[df_f['상위MD상품군명'].isin(sel_md_group)]
        if sel_md: df_f = df_f[df_f['백화점MD'].isin(sel_md)]
        if sel_brand: df_f = df_f[df_f['브랜드명'].isin(sel_brand)]
        if status_option != "전체": df_f = df_f[df_f['상품상태'] == status_option]
        df_f = df_f[(df_f['마진율'] >= sel_min_margin) & (df_f['마진율'] <= sel_max_margin)]

        total_count = len(df_f)
        st.success(f"✅ 총 **{total_count:,}**개의 상품이 추출되었습니다.")

        if total_count > 0:
            start_str = datetime.combine(start_dt, start_tm).strftime('%Y%m%d%H%M')
            end_str = datetime.combine(end_dt, end_tm).strftime('%Y%m%d%H%M')
            
            df_upload = pd.DataFrame({
                '상품번호': df_f['상품번호'],
                '매장범위': store_range[0],
                '행사시작일': start_str, '행사종료일': end_str,
                '할인유형': discount_type.split(' ')[0], '할인액': discount_value,
                '거래처분담율': vendor_share, '제휴사분담율': partner_share,
                '사용요일': day_setting, '시작시간': '0000', '종료시간': '2359', 
                '요일/시간 할인율': "" 
            })

            CHUNK_SIZE = 5000
            st.info(f"{CHUNK_SIZE:,}건 단위로 분할된 파일을 다운로드하세요.")
            
            num_chunks = (total_count // CHUNK_SIZE) + 1
            
            # Ver2의 버튼 크기를 위해 4열 배치를 유지하면서, CSS로 간격만 좁힙니다.
            cols = st.columns(4) 
            
            for idx in range(num_chunks):
                start_idx = idx * CHUNK_SIZE
                end_idx = min(start_idx + CHUNK_SIZE, total_count)
                if start_idx >= total_count: break
                
                chunk_df = df_upload.iloc[start_idx:end_idx]
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    chunk_df.to_excel(writer, index=False)
                
                with cols[idx % 4]:
                    st.download_button(
                        label=f"Part {idx+1} ({start_idx+1:,}~{end_idx:,})",
                        data=output.getvalue(),
                        file_name=f"coupon_part{idx+1}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"btn_{idx}"
                    )
        else:
            st.warning("추출된 데이터가 없습니다.")
