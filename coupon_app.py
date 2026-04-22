import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os

st.set_page_config(page_title="롯데온 Pro 쿠폰 업로드 지원(V3)", layout="wide")
st.title("🎫 롯데온 Pro 쿠폰 대량 업로드 지원 App")

# -----------------------------------------------------------------------------
# 1. 데이터 로드
# -----------------------------------------------------------------------------
DEFAULT_FILE = "product_dummy_data_30k.csv"

@st.cache_data
def load_data(file_path_or_buffer):
    try:
        if isinstance(file_path_or_buffer, str):
            return pd.read_csv(file_path_or_buffer, encoding='utf-8-sig')
        else:
            if file_path_or_buffer.name.endswith('csv'):
                return pd.read_csv(file_path_or_buffer, encoding='utf-8-sig')
            else:
                return pd.read_excel(file_path_or_buffer)
    except Exception as e:
        st.error(f"데이터 로드 오류: {e}")
        return None

st.sidebar.header("1. 데이터 소스")
df_raw = None

if os.path.exists(DEFAULT_FILE):
    st.sidebar.info(f"✅ 로컬 파일 발견: `{DEFAULT_FILE}`")
    df_raw = load_data(DEFAULT_FILE)
else:
    uploaded_file = st.sidebar.file_uploader("상품코드 데이터를 업로드하세요", type=['xlsx', 'csv'])
    if uploaded_file is not None:
        df_raw = load_data(uploaded_file)

if df_raw is not None:
    # -----------------------------------------------------------------------------
    # 2. 대상 상품 필터링
    # -----------------------------------------------------------------------------
    st.sidebar.divider()
    st.sidebar.header("2. 대상 상품 필터링")

    col_store = '상위거래처'
    col_md = '백화점MD'
    col_brand = '브랜드명'
    col_group = '상위MD상품군명'
    col_status = '상품상태'
    col_id = '상품번호'
    col_margin = '마진율'

    sel_store = st.sidebar.multiselect("점포(상위거래처)", sorted(df_raw[col_store].unique()))
    sel_md_group = st.sidebar.multiselect("상위MD상품군", sorted(df_raw[col_group].unique()))
    sel_md = st.sidebar.multiselect("백화점MD", sorted(df_raw[col_md].unique()))
    sel_brand = st.sidebar.multiselect("브랜드명", sorted(df_raw[col_brand].unique()))
    
    # [수정] 마진율 필터 추가
    if col_margin in df_raw.columns:
        st.sidebar.markdown("**마진율 범위 설정 (%)**")
        min_margin_val = int(df_raw[col_margin].min())
        max_margin_val = int(df_raw[col_margin].max())
        
        c1, c2 = st.sidebar.columns(2)
        with c1:
            sel_min_margin = st.number_input("이상 (≥)", min_value=0, max_value=100, value=min_margin_val)
        with c2:
            sel_max_margin = st.number_input("이하 (≤)", min_value=0, max_value=100, value=max_margin_val)

    status_option = st.sidebar.radio("상품 상태", ("전체", "전시", "미전시"), horizontal=True)

    # 필터 적용
    df_filtered = df_raw.copy()
    if sel_store:
        df_filtered = df_filtered[df_filtered[col_store].isin(sel_store)]
    if sel_md_group:
        df_filtered = df_filtered[df_filtered[col_group].isin(sel_md_group)]
    if sel_md:
        df_filtered = df_filtered[df_filtered[col_md].isin(sel_md)]
    if sel_brand:
        df_filtered = df_filtered[df_filtered[col_brand].isin(sel_brand)]
    if status_option != "전체":
        df_filtered = df_filtered[df_filtered[col_status] == status_option]
    
    # 마진율 필터 적용
    if col_margin in df_raw.columns:
        df_filtered = df_filtered[(df_filtered[col_margin] >= sel_min_margin) & (df_filtered[col_margin] <= sel_max_margin)]

    st.sidebar.metric("✅ 필터링된 상품 수", f"{len(df_filtered):,}개")

    # -----------------------------------------------------------------------------
    # 3. 쿠폰 정책 설정
    # -----------------------------------------------------------------------------
    st.header("쿠폰 발행 정책 설정")
    
    c1, c2, c3 = st.columns(3)
    with c1:
        store_range = st.selectbox("매장범위", ["A (전채널)", "M (본매장)", "O (외부매장)"])
        discount_type = st.selectbox("할인유형", ["10 (정률)", "30 (정액)"])
        discount_value = st.number_input("할인액/율", min_value=0, value=10)

    with c2:
        start_dt = st.date_input("시작일", datetime.now())
        start_tm = st.time_input("시작시간", datetime.strptime("00:00", "%H:%M").time())
        end_dt = st.date_input("종료일", datetime.now())
        end_tm = st.time_input("종료시간", datetime.strptime("23:59", "%H:%M").time())

    with c3:
        # [수정] 기본값 0%, 합계 검사 제거
        vendor_share = st.number_input("거래처 분담율 (%)", min_value=0, max_value=100, value=0)
        partner_share = st.number_input("제휴사 분담율 (%)", min_value=0, max_value=100, value=0)

    # -----------------------------------------------------------------------------
    # 4. 결과 생성 및 다운로드 (5,000건 단위 분할)
    # -----------------------------------------------------------------------------
    st.divider()
    
    start_str = datetime.combine(start_dt, start_tm).strftime('%Y%m%d%H%M')
    end_str = datetime.combine(end_dt, end_tm).strftime('%Y%m%d%H%M')
    
    df_upload = pd.DataFrame({
        '상품번호': df_filtered[col_id],
        '매장범위': store_range[0],
        '행사시작일': start_str,
        '행사종료일': end_str,
        '할인유형': discount_type.split(' ')[0],
        '할인액': discount_value,
        '거래처분담율': vendor_share,
        '제휴사분담율': partner_share,
        '사용요일': 'OOOOOOO',
        '시작시간': '0000',
        '종료시간': '2359',
        '요일/시간 할인율': 0
    })

    st.subheader(f"최종 업로드 데이터 (총 {len(df_upload):,}건)")
    st.dataframe(df_upload.head(100), use_container_width=True)

    # [수정] 5,000건 단위 파일 분할 로직
    CHUNK_SIZE = 5000
    total_rows = len(df_upload)
    
    if total_rows == 0:
        st.warning("다운로드할 데이터가 없습니다.")
    else:
        st.write(f"💡 업로드 시스템 안정성을 위해 최대 **{CHUNK_SIZE:,}건** 단위로 파일이 분할되어 제공됩니다.")
        
        # 다운로드 버튼을 가로로 나열
        cols = st.columns(min((total_rows // CHUNK_SIZE) + 1, 5))
        
        for i in range(0, total_rows, CHUNK_SIZE):
            chunk = df_upload.iloc[i:i + CHUNK_SIZE]
            file_part = (i // CHUNK_SIZE) + 1
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                chunk.to_excel(writer, index=False)
            
            start_row = i + 1
            end_row = min(i + CHUNK_SIZE, total_rows)
            label = f"📥 Part {file_part} 다운로드\n({start_row:,}~{end_row:,}건)"
            
            with cols[file_part % len(cols) - 1]:
                st.download_button(
                    label=label,
                    data=output.getvalue(),
                    file_name=f"coupon_upload_{datetime.now().strftime('%Y%m%d')}_part{file_part}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{file_part}"
                )