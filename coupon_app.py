import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os

# 페이지 기본 설정
st.set_page_config(page_title="롯데온 Pro 쿠폰 업로드 지원(V2)", layout="wide")

st.title("🎫 롯데온 Pro 쿠폰 대량 업로드 지원 App")
st.markdown("브랜드, 점포, 상품군, MD별 쿠폰 대량 업로드용 엑셀 파일 생성 자동화 도구입니다.")

# -----------------------------------------------------------------------------
# 1. 데이터 로드 로직 (로컬 파일 우선, 업로더 보조)
# -----------------------------------------------------------------------------
DEFAULT_FILE = "product_dummy_data_30k.csv"

@st.cache_data
def load_data(file_path_or_buffer):
    try:
        if isinstance(file_path_or_buffer, str):
            # 로컬 파일 읽기 (UTF-8-SIG는 한글 깨짐 방지용)
            return pd.read_csv(file_path_or_buffer, encoding='utf-8-sig')
        else:
            # 업로드된 파일 읽기
            if file_path_or_buffer.name.endswith('csv'):
                return pd.read_csv(file_path_or_buffer, encoding='utf-8-sig')
            else:
                return pd.read_excel(file_path_or_buffer)
    except Exception as e:
        st.error(f"데이터 로드 중 오류 발생: {e}")
        return None

# 사이드바에서 데이터 소스 확인
st.sidebar.header("1. 데이터 소스")

df_raw = None

# 1순위: 로컬에 지정된 이름의 파일이 있는지 확인
if os.path.exists(DEFAULT_FILE):
    st.sidebar.info(f"✅ 로컬 파일 발견: `{DEFAULT_FILE}`")
    df_raw = load_data(DEFAULT_FILE)
else:
    st.sidebar.warning(f"⚠️ `{DEFAULT_FILE}` 파일을 찾을 수 없습니다.")

# 2순위: 사용자가 직접 파일을 업로드할 수 있는 옵션 유지
uploaded_file = st.sidebar.file_uploader("또는 다른 상품코드 데이터를 업로드하세요", type=['xlsx', 'csv'])
if uploaded_file is not None:
    df_raw = load_data(uploaded_file)

if df_raw is not None:
    st.sidebar.success(f"총 {len(df_raw):,}개의 상품 데이터를 로드했습니다.")

    # -----------------------------------------------------------------------------
    # 2. Sidebar: 상세 필터링
    # -----------------------------------------------------------------------------
    st.sidebar.divider()
    st.sidebar.header("2. 대상 상품 필터링")

    # 컬럼명 매핑 (더미 데이터 구조 기준)
    col_store = '상위거래처'
    col_md = '백화점MD'
    col_brand = '브랜드명'
    col_group = '상위MD상품군명'
    col_status = '상품상태'
    col_id = '상품번호'

    # 필터 선택 UI
    sel_store = st.sidebar.multiselect("점포(상위거래처) 선택", sorted(df_raw[col_store].unique()))
    sel_md_group = st.sidebar.multiselect("상위MD상품군 선택", sorted(df_raw[col_group].unique()))
    sel_md = st.sidebar.multiselect("백화점MD 선택", sorted(df_raw[col_md].unique()))
    sel_brand = st.sidebar.multiselect("브랜드명 선택", sorted(df_raw[col_brand].unique()))
    
    status_option = st.sidebar.radio(
        "상품 상태 필터링",
        ("전체", "전시", "미전시"),
        horizontal=True
    )

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

    st.sidebar.metric("필터링된 상품 수", f"{len(df_filtered):,}개")

    # -----------------------------------------------------------------------------
    # 3. Main Dashboard: 쿠폰 정책 설정
    # -----------------------------------------------------------------------------
    st.header("쿠폰 발행 정책 설정")
    
    c1, c2, c3 = st.columns(3)
    with c1:
        st.subheader("📍 범위 및 유형")
        store_range = st.selectbox("매장범위", ["A (전채널)", "M (본매장)", "O (외부매장)"])
        discount_type = st.selectbox("할인유형", ["10 (정률)", "30 (정액)"])
        discount_value = st.number_input("할인액/율", min_value=0, value=10)

    with c2:
        st.subheader("📅 행사 기간")
        start_dt = st.date_input("시작일", datetime.now())
        start_tm = st.time_input("시작시간", datetime.strptime("00:00", "%H:%M").time())
        end_dt = st.date_input("종료일", datetime.now())
        end_tm = st.time_input("종료시간", datetime.strptime("23:59", "%H:%M").time())

    with c3:
        st.subheader("💰 비용 분담 (%)")
        vendor_share = st.number_input("거래처 분담율", 0, 100, 50)
        partner_share = st.number_input("제휴사 분담율", 0, 100, 50)
        
        total_share = vendor_share + partner_share
        if total_share == 100:
            st.success(f"합계: {total_share}% (정상)")
        else:
            st.error(f"합계: {total_share}% (100%가 되어야 합니다)")

    # -----------------------------------------------------------------------------
    # 4. 결과 생성 및 다운로드
    # -----------------------------------------------------------------------------
    st.divider()
    if st.button("🚀 업로드용 데이터 생성", use_container_width=True, disabled=(total_share != 100)):
        start_str = datetime.combine(start_dt, start_tm).strftime('%Y%m%d%H%M')
        end_str = datetime.combine(end_dt, end_tm).strftime('%Y%m%d%H%M')
        
        # 엑셀 양식에 맞춘 데이터프레임 구성
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

        st.subheader("최종 파일 미리보기 (상위 100건)")
        st.dataframe(df_upload.head(100), use_container_width=True)

        # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_upload.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 coupon_upload.xlsx 다운로드",
            data=output.getvalue(),
            file_name=f"coupon_upload_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("사이드바에 `product_dummy_data_30k.csv` 파일이 있거나 데이터를 업로드해야 앱이 작동합니다.")