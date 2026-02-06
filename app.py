import streamlit as st
from pathlib import Path
import pandas as pd

from inventory_search import load_inventory, load_inventory_from_bytes, search_inventory, summarize_inventory, lot_breakdown
from inventory_search import (
    build_summary_export_multi,
    COL_COLOR,
    COL_COLOR_HEX,
    COL_NAME,
    COL_P,
    COL_T,
    COL_POWER,
    COL_TONE,
    COL_CYL,
    COL_AXIS,
    COL_ADD,
)

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

header_left, header_right = st.columns([1, 6])
with header_left:
    st.subheader("통합 검색")
with header_right:
    toric_mf = st.checkbox("Toric+M/F", value=False)
col1, col2, col3, col4, col5, col6, col7 = st.columns([2, 1, 1, 1, 1, 1, 1])
with col1:
    query = st.text_input("코드/품명", placeholder="예: T4556, P4050, NewFusion")
with col2:
    color_query = st.text_input("컬러(컬러코드)", placeholder="예: Blue, BL01")
with col3:
    tone_query = st.text_input("톤수", placeholder="예: 2")
with col4:
    power_query = st.text_input("파워", placeholder="예: -02.50")
with col5:
    cyl_query = st.text_input("CYL", placeholder="예: -01.25")
with col6:
    axis_query = st.text_input("AXIS", placeholder="예: 90")
with col7:
    add_query = st.text_input("ADD", placeholder="예: +1.00")


missing_cols: list[str] = []


def _filter_contains(df: pd.DataFrame, col: str, value: str) -> pd.DataFrame:
    if not value:
        return df
    if col not in df.columns:
        missing_cols.append(col)
        return df
    value_upper = value.strip().upper()
    series = df[col].fillna("").astype(str).str.strip().str.upper()
    return df[series.str.contains(value_upper, na=False)]


def _filter_numeric_equal(df: pd.DataFrame, col: str, value: str, decimals: int = 2) -> pd.DataFrame:
    if not value:
        return df
    if col not in df.columns:
        missing_cols.append(col)
        return df
    try:
        target = round(float(str(value).strip()), decimals)
    except Exception:
        return df
    series = pd.to_numeric(df[col], errors="coerce").round(decimals)
    return df[series == target]


def _filter_int_equal(df: pd.DataFrame, col: str, value: str) -> pd.DataFrame:
    if not value:
        return df
    if col not in df.columns:
        missing_cols.append(col)
        return df
    try:
        target = int(str(value).strip())
    except Exception:
        return df
    series = pd.to_numeric(df[col], errors="coerce").round(0)
    return df[series == target]


def _filter_color(df: pd.DataFrame, value: str) -> pd.DataFrame:
    if not value:
        return df
    value_upper = value.strip().upper().replace(" ", "")
    masks = []
    if COL_COLOR in df.columns:
        s = df[COL_COLOR].fillna("").astype(str).str.upper().str.replace(r"\s+", "", regex=True)
        masks.append(s.str.contains(value_upper, na=False))
    if "컬러코드" in df.columns:
        s = df["컬러코드"].fillna("").astype(str).str.upper().str.replace(r"\s+", "", regex=True)
        masks.append(s.str.contains(value_upper, na=False))
    if not masks:
        missing_cols.append("컬러/컬러코드")
        return df
    mask = masks[0]
    for m in masks[1:]:
        mask = mask | m
    return df[mask]

has_filters = any(
    [
        query.strip(),
        color_query.strip(),
        tone_query.strip(),
        power_query.strip(),
        cyl_query.strip(),
        axis_query.strip(),
        add_query.strip(),
    ]
)

if has_filters:
    filtered_df = df
    filtered_df = _filter_color(filtered_df, color_query)
    filtered_df = _filter_contains(filtered_df, COL_TONE, tone_query)
    filtered_df = _filter_numeric_equal(filtered_df, COL_POWER, power_query, decimals=2)
    filtered_df = _filter_numeric_equal(filtered_df, COL_CYL, cyl_query, decimals=2)
    filtered_df = _filter_int_equal(filtered_df, COL_AXIS, axis_query)
    filtered_df = _filter_numeric_equal(filtered_df, COL_ADD, add_query, decimals=2)

    if missing_cols:
        missing_label = ", ".join(sorted(set(missing_cols)))
        st.warning(f"엑셀에 다음 컬럼이 없어 해당 조건은 무시했습니다: {missing_label}")

    if query.strip():
        result = search_inventory(filtered_df, query)
    else:
        result = summarize_inventory(filtered_df)
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
                use_toric=toric_mf,
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
    st.info("검색 조건을 입력하면 결과가 표시됩니다.")
