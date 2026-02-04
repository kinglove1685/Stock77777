# inventory_search.py
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles.colors import COLOR_INDEX

SHEET_NAME = "재고장"

# Column names in the source sheet
COL_P = "P코드"
COL_T = "T코드"
COL_U = "U코드"
COL_ITEM = "품목코드"
COL_NAME = "품명"
COL_STOCK = "재고"
COL_POWER = "파워"
COL_COLOR = "컬러"
COL_TONE = "톤수"
COL_LOTNO = "LOTNO"
COL_COLOR_HEX = "컬러_색상"
SUMMARY_SHEET = "SUMMARY"

DEFAULT_SEARCH_COLS = [COL_P, COL_T, COL_U, COL_NAME, COL_ITEM, COL_COLOR, COL_TONE]

CODE_PATTERN = re.compile(r"^[PTU]\d+$", re.IGNORECASE)
CODE_COLOR_PATTERN = re.compile(r"^([PTU]\d+)(.+)$", re.IGNORECASE)


def _normalize_series(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip()

def _compact_upper(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.upper().str.replace(r"\s+", "", regex=True)


def _excel_color_to_hex(cell) -> str:
    fill = cell.fill
    if fill is None or fill.patternType is None or fill.fgColor is None:
        return ""
    color = fill.fgColor
    if color.type == "rgb" and color.rgb:
        rgb = color.rgb
        if len(rgb) == 8:
            rgb = rgb[2:]
        return f"#{rgb}"
    if color.type == "indexed" and color.indexed is not None:
        try:
            rgb = COLOR_INDEX[color.indexed]
            if rgb:
                if len(rgb) == 8:
                    rgb = rgb[2:]
                return f"#{rgb}"
        except Exception:
            return ""
    return ""


def _load_color_column_from_workbook(wb, data_len: int) -> list[str]:
    # Excel column E -> 5 (1-based)
    ws = wb[SHEET_NAME]
    colors = []
    for row_idx in range(2, data_len + 2):
        cell = ws.cell(row=row_idx, column=5)
        colors.append(_excel_color_to_hex(cell))
    return colors


def load_inventory(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SHEET_NAME)

    # Normalize key columns to strings
    for col in [COL_P, COL_T, COL_U, COL_ITEM, COL_NAME, COL_COLOR, COL_TONE, COL_LOTNO]:
        if col in df.columns:
            df[col] = _normalize_series(df[col])

    # Stock: NaN -> 0
    if COL_STOCK in df.columns:
        df[COL_STOCK] = pd.to_numeric(df[COL_STOCK], errors="coerce").fillna(0)

    # Load fill color from Excel column E (컬러)
    try:
        wb = load_workbook(path, data_only=True)
        if SHEET_NAME in wb.sheetnames:
            df[COL_COLOR_HEX] = _load_color_column_from_workbook(wb, len(df))
    except Exception:
        df[COL_COLOR_HEX] = ""

    return df


def load_inventory_from_bytes(data: bytes) -> pd.DataFrame:
    from io import BytesIO

    bio = BytesIO(data)
    df = pd.read_excel(bio, sheet_name=SHEET_NAME)

    for col in [COL_P, COL_T, COL_U, COL_ITEM, COL_NAME, COL_COLOR, COL_TONE, COL_LOTNO]:
        if col in df.columns:
            df[col] = _normalize_series(df[col])

    if COL_STOCK in df.columns:
        df[COL_STOCK] = pd.to_numeric(df[COL_STOCK], errors="coerce").fillna(0)

    try:
        bio.seek(0)
        wb = load_workbook(bio, data_only=True)
        if SHEET_NAME in wb.sheetnames:
            df[COL_COLOR_HEX] = _load_color_column_from_workbook(wb, len(df))
    except Exception:
        df[COL_COLOR_HEX] = ""

    return df


def _open_workbook_from_source(path: Path | None, data: bytes | None):
    from io import BytesIO

    if data is not None:
        bio = BytesIO(data)
        return load_workbook(bio, data_only=True, read_only=True)
    if path is None:
        raise ValueError("Either path or data must be provided")
    return load_workbook(path, data_only=True, read_only=True)


def _find_summary_column(ws, pcode: str, color: str, name: str) -> int | None:
    max_col = ws.max_column
    p_norm = str(pcode).strip().upper()
    c_norm = str(color).strip().upper() if color else ""
    n_norm = str(name).strip().upper() if name else ""

    # Match by P코드 + 컬러 (if provided)
    for c in range(3, max_col + 1):
        p = ws.cell(row=3, column=c).value
        if p is None:
            continue
        if str(p).strip().upper() != p_norm:
            continue
        if c_norm:
            col_val = ws.cell(row=4, column=c).value
            if str(col_val).strip().upper() != c_norm:
                continue
        return c

    # If 컬러가 없거나 매칭 실패했으면 P코드만으로 첫 컬럼 선택
    if p_norm:
        for c in range(3, max_col + 1):
            p = ws.cell(row=3, column=c).value
            if p is None:
                continue
            if str(p).strip().upper() == p_norm:
                return c

    # Fallback: match by 품명 on row 2
    if n_norm:
        for c in range(3, max_col + 1):
            v = ws.cell(row=2, column=c).value
            if v is None:
                continue
            if str(v).strip().upper() == n_norm:
                return c

    return None


def build_summary_export(
    pcode: str,
    color: str,
    name: str,
    source_path: Path | None = None,
    source_bytes: bytes | None = None,
) -> tuple[bytes, bool]:
    wb = _open_workbook_from_source(source_path, source_bytes)
    if SUMMARY_SHEET not in wb.sheetnames:
        raise ValueError("SUMMARY sheet not found")
    ws = wb[SUMMARY_SHEET]

    target_col = _find_summary_column(ws, pcode, color, name)

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "EXPORT"

    summary_name = None
    if target_col is not None:
        summary_name = ws.cell(row=2, column=target_col).value

    out_ws["C2"] = summary_name if summary_name else name
    out_ws["C3"] = pcode
    out_ws["C4"] = color or ""
    out_ws["B4"] = "파워"

    row = 5
    total = 0.0
    for r in range(5, ws.max_row + 1):
        power = ws.cell(row=r, column=2).value
        if power is None or str(power).strip() == "":
            break
        qty = 0
        if target_col is not None:
            val = ws.cell(row=r, column=target_col).value
            qty = 0 if val is None else val

        out_ws.cell(row=row, column=2, value=power)
        out_ws.cell(row=row, column=3, value=qty)

        try:
            total += float(qty)
        except Exception:
            pass
        row += 1

    out_ws.cell(row=row, column=2, value="합계")
    out_ws.cell(row=row, column=3, value=total)

    from io import BytesIO

    bio = BytesIO()
    out_wb.save(bio)
    return bio.getvalue(), target_col is not None


def build_summary_export_multi(
    rows: list[dict],
    source_path: Path | None = None,
    source_bytes: bytes | None = None,
) -> tuple[bytes, int]:
    wb = _open_workbook_from_source(source_path, source_bytes)
    if SUMMARY_SHEET not in wb.sheetnames:
        raise ValueError("SUMMARY sheet not found")
    ws = wb[SUMMARY_SHEET]

    # Collect power rows from SUMMARY (column B, starting row 5)
    powers = []
    r = 5
    while True:
        v = ws.cell(row=r, column=2).value
        if v is None or str(v).strip() == "":
            break
        if str(v).strip() in ("합계", "총합계"):
            break
        powers.append(v)
        r += 1

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "EXPORT"

    out_ws["B4"] = "컬러"
    out_ws["B5"] = "톤수"
    out_ws["B6"] = "파워"

    # Build product -> power qty map from search result rows
    product_keys = []
    power_map = {}
    for row in rows:
        pcode = str(row.get(COL_P, "")).strip()
        color = str(row.get(COL_COLOR, "")).strip()
        name = str(row.get(COL_NAME, "")).strip()
        tone = str(row.get(COL_TONE, "")).strip()
        power = str(row.get(COL_POWER, "")).strip()
        qty = row.get(COL_STOCK, 0)
        key = (pcode, color, tone, name)
        if key not in power_map:
            power_map[key] = {}
            product_keys.append(key)
        p_norm = power.strip().upper()
        try:
            qty_val = float(qty)
        except Exception:
            qty_val = 0.0
        power_map[key][p_norm] = power_map[key].get(p_norm, 0.0) + qty_val

    not_found = 0
    for idx, key in enumerate(product_keys):
        pcode, color, tone, name = key

        out_col = 3 + idx
        target_col = _find_summary_column(ws, pcode, color, name)
        if target_col is None:
            not_found += 1

        summary_name = None
        if target_col is not None:
            summary_name = ws.cell(row=2, column=target_col).value

        out_ws.cell(row=2, column=out_col, value=summary_name if summary_name else name)
        out_ws.cell(row=3, column=out_col, value=pcode)
        out_ws.cell(row=4, column=out_col, value=color or "")
        out_ws.cell(row=5, column=out_col, value=tone or "")

        total = 0.0
        local_map = power_map.get(key, {})
        for i, power in enumerate(powers):
            row = 6 + i
            p_norm = str(power).strip().upper()
            qty = local_map.get(p_norm, 0.0)
            out_ws.cell(row=row, column=2, value=power)
            out_ws.cell(row=row, column=out_col, value=qty)
            try:
                total += float(qty)
            except Exception:
                pass

        out_ws.cell(row=6 + len(powers), column=2, value="합계")
        out_ws.cell(row=6 + len(powers), column=out_col, value=total)

    from io import BytesIO

    bio = BytesIO()
    out_wb.save(bio)
    return bio.getvalue(), not_found


def _term_mask(df: pd.DataFrame, term: str) -> pd.Series:
    term = term.strip()
    if not term:
        return pd.Series(False, index=df.index)

    term_upper = term.upper()
    code_color_match = CODE_COLOR_PATTERN.match(term_upper)
    is_code_like = bool(CODE_PATTERN.match(term_upper))

    masks = []

    if code_color_match and not is_code_like:
        code = code_color_match.group(1)
        color_term = code_color_match.group(2).strip()
        tone_term = ""
        tone_match = re.match(r"^(.*?)(\d+)$", color_term)
        if tone_match:
            color_term = tone_match.group(1).strip()
            tone_term = tone_match.group(2).strip()

        if COL_P in df.columns:
            masks.append(df[COL_P].str.upper() == code)
        if COL_T in df.columns:
            masks.append(df[COL_T].str.upper() == code)
        if COL_U in df.columns:
            masks.append(df[COL_U].str.upper() == code)

        color_masks = []
        if COL_COLOR in df.columns:
            term_compact = re.sub(r"\s+", "", color_term).upper()
            color_masks.append(_compact_upper(df[COL_COLOR]).str.contains(term_compact, na=False))
        if "컬러코드" in df.columns:
            term_compact = re.sub(r"\s+", "", color_term).upper()
            color_masks.append(_compact_upper(df["컬러코드"]).str.contains(term_compact, na=False))

        if tone_term and COL_TONE in df.columns:
            tone_mask = df[COL_TONE].str.upper() == tone_term.upper()
        else:
            tone_mask = None

        if color_masks:
            color_mask = color_masks[0]
            for m in color_masks[1:]:
                color_mask = color_mask | m

            if masks:
                code_mask = masks[0]
                for m in masks[1:]:
                    code_mask = code_mask | m
                combined = code_mask & color_mask
                if tone_mask is not None:
                    combined = combined & tone_mask
                return combined
            combined = color_mask
            if tone_mask is not None:
                combined = combined & tone_mask
            return combined

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

    # Group by P/T + color + tone + power and keep a representative name
    group_cols = []
    if COL_P in df.columns:
        group_cols.append(COL_P)
    if COL_T in df.columns:
        group_cols.append(COL_T)
    if COL_COLOR in df.columns:
        group_cols.append(COL_COLOR)
    if COL_TONE in df.columns:
        group_cols.append(COL_TONE)
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
                "LOTNO 합계수량": (COL_LOTNO, lambda s: s.notna().sum()) if COL_LOTNO in df.columns else (COL_STOCK, "count"),
                COL_COLOR_HEX: (COL_COLOR_HEX, "first") if COL_COLOR_HEX in df.columns else (COL_P, "first"),
            }
        )
        .reset_index()
        .sort_values(COL_STOCK, ascending=False)
    )

    if COL_STOCK in result.columns:
        result[COL_STOCK] = pd.to_numeric(result[COL_STOCK], errors="coerce").fillna(0).round(0).astype(int)

    # Power formatting: -00.00, -05.25, -04.75
    if COL_POWER in result.columns:
        def _fmt_power(v):
            try:
                fv = float(v)
            except Exception:
                return v
            sign = "-" if fv <= 0 else "+"
            return f"{sign}{abs(fv):05.2f}"

        result[COL_POWER] = result[COL_POWER].apply(_fmt_power)

    # Column order: P, T, 컬러, 톤수, 파워, 품명, 재고, LOTNO 합계수량
    preferred = [COL_P, COL_T, COL_COLOR, COL_TONE, COL_POWER, COL_NAME, COL_STOCK, "LOTNO 합계수량", COL_COLOR_HEX]
    cols = [c for c in preferred if c in result.columns]
    if cols:
        result = result[cols]

    return result


def lot_breakdown(df: pd.DataFrame, row: dict) -> pd.DataFrame:
    # Filter original rows to match selected summary row
    filters = []
    if COL_P in df.columns and row.get(COL_P) is not None:
        filters.append(df[COL_P] == str(row.get(COL_P, "")).strip())
    if COL_T in df.columns and row.get(COL_T) is not None:
        filters.append(df[COL_T] == str(row.get(COL_T, "")).strip())
    if COL_COLOR in df.columns and row.get(COL_COLOR) is not None:
        filters.append(df[COL_COLOR] == str(row.get(COL_COLOR, "")).strip())
    if COL_TONE in df.columns and row.get(COL_TONE) is not None:
        filters.append(df[COL_TONE] == str(row.get(COL_TONE, "")).strip())
    if COL_POWER in df.columns and row.get(COL_POWER) is not None:
        try:
            power_val = float(str(row.get(COL_POWER, "")).strip())
            power_series = pd.to_numeric(df[COL_POWER], errors="coerce").round(2)
            filters.append(power_series == round(power_val, 2))
        except Exception:
            filters.append(df[COL_POWER].astype(str) == str(row.get(COL_POWER, "")).strip())

    if not filters:
        return df.head(0)

    mask = filters[0]
    for f in filters[1:]:
        mask = mask & f

    filtered = df[mask]
    if COL_LOTNO not in filtered.columns:
        return filtered.head(0)

    lot_df = (
        filtered.groupby(COL_LOTNO, dropna=False)[COL_STOCK]
        .sum()
        .reset_index()
        .sort_values(COL_STOCK, ascending=False)
    )
    if COL_STOCK in lot_df.columns:
        lot_df[COL_STOCK] = pd.to_numeric(lot_df[COL_STOCK], errors="coerce").fillna(0).round(0).astype(int)

    return lot_df


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
