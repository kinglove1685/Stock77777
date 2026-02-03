import streamlit as st
from pathlib import Path
import pandas as pd

from inventory_search import load_inventory, search_inventory

DEFAULT_PATH = Path(__file__).parent / "재고관련 프로그램제작.xlsx"

st.set_page_config(page_title="재고 검색", layout="wide")

st.title("재고 검색 대시보드")

with st.sidebar:
    st.header("데이터")
    use_upload = st.checkbox("엑셀 업로드 사용", value=False)
    uploaded_file = None
    if use_upload:
        uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx", "xlsm", "xls"])
    st.caption("기본 경로: " + str(DEFAULT_PATH))

@st.cache_data(show_spinner=False)
def load_from_path(path: Path) -> pd.DataFrame:
    return load_inventory(path)

@st.cache_data(show_spinner=False)
def load_from_bytes(data: bytes) -> pd.DataFrame:
    # Use BytesIO for uploaded files
    from io import BytesIO
    return pd.read_excel(BytesIO(data), sheet_name="재고장")

# Load data
if use_upload and uploaded_file is not None:
    df = load_from_bytes(uploaded_file.getvalue())
    # Normalize columns to match inventory_search expectations
    for col in ["P코드", "T코드", "U코드", "품목코드", "품명"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    if "재고" in df.columns:
        df["재고"] = pd.to_numeric(df["재고"], errors="coerce").fillna(0)
else:
    if not DEFAULT_PATH.exists():
        st.error(f"기본 엑셀 파일이 없습니다: {DEFAULT_PATH}")
        st.stop()
    df = load_from_path(DEFAULT_PATH)

st.subheader("통합 검색")
query = st.text_input("코드/품명 키워드 입력", placeholder="예: T4556, P4050, NewFusion")

if query.strip():
    result = search_inventory(df, query)
    if result.empty:
        st.info("검색 결과가 없습니다.")
    else:
        st.dataframe(result, use_container_width=True)
else:
    st.info("검색어를 입력하면 결과가 표시됩니다.")
