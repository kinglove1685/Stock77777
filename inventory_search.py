# inventory_search.py
import re
import sys
from pathlib import Path

import pandas as pd

SHEET_NAME = "재고장"

# Column names in the source sheet
COL_P = "P코드"
COL_T = "T코드"
COL_U = "U코드"
COL_ITEM = "품목코드"
COL_NAME = "품명"
COL_STOCK = "재고"
COL_POWER = "파워"

DEFAULT_SEARCH_COLS = [COL_P, COL_T, COL_U, COL_NAME, COL_ITEM]

CODE_PATTERN = re.compile(r"^[PTU]\d+$", re.IGNORECASE)


def _normalize_series(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip()


def load_inventory(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SHEET_NAME)

    # Normalize key columns to strings
    for col in [COL_P, COL_T, COL_U, COL_ITEM, COL_NAME]:
        if col in df.columns:
            df[col] = _normalize_series(df[col])

    # Stock: NaN -> 0
    if COL_STOCK in df.columns:
        df[COL_STOCK] = pd.to_numeric(df[COL_STOCK], errors="coerce").fillna(0)

    return df


def _term_mask(df: pd.DataFrame, term: str) -> pd.Series:
    term = term.strip()
    if not term:
        return pd.Series(False, index=df.index)

    is_code_like = bool(CODE_PATTERN.match(term))
    term_upper = term.upper()

    masks = []

    if is_code_like:
        # Exact match for code fields
        if COL_P in df.columns:
            masks.append(df[COL_P].str.upper() == term_upper)
        if COL_T in df.columns:
            masks.append(df[COL_T].str.upper() == term_upper)
        if COL_U in df.columns:
            masks.append(df[COL_U].str.upper() == term_upper)
        # Allow partial match on item code for convenience
        if COL_ITEM in df.columns:
            masks.append(df[COL_ITEM].str.upper().str.contains(term_upper, na=False))
    else:
        # Partial match across name + codes
        for col in DEFAULT_SEARCH_COLS:
            if col in df.columns:
                masks.append(df[col].str.upper().str.contains(term_upper, na=False))

    if not masks:
        return pd.Series(False, index=df.index)

    mask = masks[0]
    for m in masks[1:]:
        mask = mask | m
    return mask


def search_inventory(df: pd.DataFrame, query: str) -> pd.DataFrame:
    # Split by spaces or commas
    terms = [t for t in re.split(r"[\s,]+", query) if t.strip()]
    if not terms:
        return df.head(0)

    mask = pd.Series(False, index=df.index)
    for term in terms:
        mask = mask | _term_mask(df, term)

    filtered = df[mask]

    # Group by P/T codes + power and keep a representative name
    group_cols = []
    if COL_P in df.columns:
        group_cols.append(COL_P)
    if COL_T in df.columns:
        group_cols.append(COL_T)
    if COL_POWER in df.columns:
        group_cols.append(COL_POWER)

    if not group_cols:
        return filtered.head(0)

    result = (
        filtered.groupby(group_cols, dropna=False)
        .agg(
            **{
                COL_NAME: (COL_NAME, "first") if COL_NAME in df.columns else (COL_P, "first"),
                COL_STOCK: (COL_STOCK, "sum"),
                "rows": (COL_STOCK, "count"),
            }
        )
        .reset_index()
        .sort_values(COL_STOCK, ascending=False)
    )

    return result


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python inventory_search.py <query> [excel_path]")
        return 1

    query = sys.argv[1]
    path = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).parent / "재고관련 프로그램제작.xlsx"

    if not path.exists():
        print(f"File not found: {path}")
        return 1

    df = load_inventory(path)
    result = search_inventory(df, query)

    if result.empty:
        print("No matches")
        return 0

    # Print a readable table
    print(result.to_string(index=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
