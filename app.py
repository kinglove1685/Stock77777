import streamlit as st
from pathlib import Path
import pandas as pd

from inventory_search import load_inventory, load_inventory_from_bytes, search_inventory, lot_breakdown
from inventory_search import build_summary_export_multi, COL_COLOR, COL_COLOR_HEX, COL_NAME, COL_P, COL_T, COL_POWER

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
    return load_inventory_from_bytes(data)

# Load data
source_bytes = None
source_path = None
if use_upload and uploaded_file is not None:
    source_bytes = uploaded_file.getvalue()
    df = load_from_bytes(source_bytes)
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
    source_path = DEFAULT_PATH

st.subheader("통합 검색")
query = st.text_input("코드/품명 키워드 입력", placeholder="예: T4556, P4050, NewFusion")

if query.strip():
    result = search_inventory(df, query)
    if result.empty:
        st.info("검색 결과가 없습니다.")
    else:
        display_df = result.copy()
        if COL_COLOR_HEX in display_df.columns:
            display_df = display_df.drop(columns=[COL_COLOR_HEX])

        left_col, right_col = st.columns([3, 1])
        with left_col:
            table = st.dataframe(
                display_df,
                use_container_width=True,
                on_select="rerun",
                selection_mode="single-row",
                height=460,
            )

        with right_col:
            if table and table.selection and table.selection.rows:
                sel_idx = table.selection.rows[0]
                selected_row = result.reset_index(drop=True).loc[sel_idx].to_dict()
                lot_df = lot_breakdown(df, selected_row)
                st.caption("LOT 상세")
                if not lot_df.empty:
                    st.dataframe(lot_df, use_container_width=True, height=320)
                else:
                    st.caption("LOT 정보 없음")

        export_df = result.reset_index(drop=True)
        if not export_df.empty:
            export_bytes, not_found = build_summary_export_multi(
                rows=export_df.to_dict(orient="records"),
                source_path=source_path,
                source_bytes=source_bytes,
            )
            if not_found:
                st.warning(f"SUMMARY 시트에서 {not_found}개 품목을 찾지 못해 수량이 0으로 출력될 수 있습니다.")

            file_name = "SEARCH_RESULT_SUMMARY_EXPORT.xlsx"

            left, _ = st.columns([1, 5])
            with left:
                st.download_button(
                    "엑셀로 내보내기",
                    data=export_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
else:
    st.info("검색어를 입력하면 결과가 표시됩니다.")
