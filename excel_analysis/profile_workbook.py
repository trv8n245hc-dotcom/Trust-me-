#!/usr/bin/env python3
"""
profile_workbook.py - generic metadata generator for any .xlsx workbook.

Produces three tables:
  * sheet_meta  - one row per sheet  (size, used range, formulas, merges, header rows,
                                      period coverage, data-type mix, ...)
  * column_meta - one row per sheet x column (header text, period, data-type counts,
                                      numeric stats, blanks, % share of rates, ...)
  * row_meta    - one row per sheet x non-empty row (section, brand, item label,
                                      metric type, numeric stats, first/last values, ...)

Usage (stand-alone):
    python profile_workbook.py input.xlsx [output_metadata.xlsx]

The functions are also imported by build_report.py so the same metadata is
written into the analysis workbook.
"""
from __future__ import annotations

import re
import statistics as st
import sys
from collections import Counter
from dataclasses import dataclass, field

import openpyxl
from openpyxl.utils import get_column_letter as col_letter

# --------------------------------------------------------------------------------------
# Domain hints (kept generic: they only drive the *labels* "brand" / "metric type")
# --------------------------------------------------------------------------------------
BRAND_ALIASES = {
    "budget": "Budget",
    "first for women": "First For Women", "ffw": "First For Women",
    "dialdirect": "Dialdirect", "dial": "Dialdirect",
    "auto & general": "Auto & General", "a&g": "Auto & General",
    "virseker": "Virseker",
    "tih multibrand": "TIH Multibrand",
}
RATE_WORDS = ("%", "conversion", "closing", "effectiveness", "cxd", "ratio",
              "distribution", "dedupe", "spoken", "tat ")
PER_DAY_WORDS = ("per day", "pd")
AVG_WORDS = ("avg ", "average", "per processed", "per contact")
PERIOD_RE = re.compile(r"^(20\d{2})(0[1-9]|1[0-2])$")
FISCAL_RE = re.compile(r"^F\d{2}\b")


def norm_brand(text) -> str | None:
    if not isinstance(text, str):
        return None
    t = text.strip().lower()
    t = re.sub(r"\s*-\s*brand total$", "", t)
    return BRAND_ALIASES.get(t)


def metric_type(label: str | None) -> str:
    if not label:
        return "unknown"
    l = label.lower()
    if any(w in l for w in AVG_WORDS):
        return "average"
    if any(w in l for w in PER_DAY_WORDS):
        return "per-day volume"
    if any(w in l for w in RATE_WORDS):
        return "rate / ratio"
    if "spv" in l or "vat" in l:
        return "currency"
    return "volume / count"


ERROR_VALUES = {"#DIV/0!", "#VALUE!", "#REF!", "#N/A", "#NAME?", "#NUM!", "#NULL!"}


def is_err(v) -> bool:
    return isinstance(v, str) and v.strip() in ERROR_VALUES


def is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def is_period(v) -> bool:
    return (is_num(v) and float(v).is_integer() and PERIOD_RE.match(str(int(v))) is not None)


# --------------------------------------------------------------------------------------
@dataclass
class SheetGrid:
    name: str
    rows: list            # list of tuples (values)
    max_row: int
    max_col: int
    formulas: int = 0
    merged: int = 0
    formula_cells: list = field(default_factory=list)


def load_grids(path: str) -> list[SheetGrid]:
    """Read every sheet once for values and once (cheaply) for formulas/merges."""
    wb_v = openpyxl.load_workbook(path, read_only=True, data_only=True)
    grids = []
    for ws in wb_v.worksheets:
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
        mc = max((len(r) for r in rows), default=0)
        grids.append(SheetGrid(ws.title, rows, len(rows), mc))
    wb_v.close()
    # formulas & merges need the non-read-only loader
    wb_f = openpyxl.load_workbook(path, read_only=False, data_only=False)
    for g in grids:
        ws = wb_f[g.name]
        g.merged = len(ws.merged_cells.ranges)
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    g.formulas += 1
                    if len(g.formula_cells) < 20:
                        g.formula_cells.append(c.coordinate)
    wb_f.close()
    return grids


def cell(g: SheetGrid, r: int, c: int):
    """1-based accessor."""
    if r - 1 >= len(g.rows):
        return None
    row = g.rows[r - 1]
    return row[c - 1] if c - 1 < len(row) else None


# --------------------------------------------------------------------------------------
def detect_header_rows(g: SheetGrid, scan: int = 12) -> dict:
    """Find period rows (YYYYMM), working-day rows and caption rows in the top block."""
    info = {"period_rows": [], "working_day_row": None, "caption_rows": []}
    for r in range(1, min(scan, g.max_row) + 1):
        vals = [v for v in g.rows[r - 1] if v not in (None, "")]
        if not vals:
            continue
        n_period = sum(is_period(v) for v in vals)
        n_text = sum(isinstance(v, str) for v in vals)
        n_small_int = sum(is_num(v) and float(v).is_integer() and 1 <= v <= 400 for v in vals)
        if n_period >= 6 or n_period >= max(3, 0.5 * len(vals)):
            info["period_rows"].append(r)
        elif n_small_int >= max(5, 0.7 * len(vals)):
            info["working_day_row"] = info["working_day_row"] or r
        elif n_text >= max(3, 0.5 * len(vals)):
            info["caption_rows"].append(r)
    return info


def column_headers(g: SheetGrid, hdr: dict) -> dict[int, str]:
    """Best header text per column: caption row text > period(from-to) > blank."""
    heads = {}
    pr = hdr["period_rows"]
    for c in range(1, g.max_col + 1):
        parts = []
        for r in hdr["caption_rows"]:
            v = cell(g, r, c)
            if isinstance(v, str) and v.strip():
                parts.append(v.strip())
        if not parts and pr:
            vals = [cell(g, r, c) for r in pr]
            vals = [str(int(v)) for v in vals if is_period(v)]
            if vals:
                parts.append(vals[0] if len(set(vals)) == 1 else f"{vals[0]}-{vals[-1]}")
        if parts:
            heads[c] = " / ".join(dict.fromkeys(parts))
    return heads


def exec_layout_headers(g: SheetGrid) -> dict[int, str]:
    """Exec-style sheets: metric group in row 2, measure (Actual/Target...) in row 3."""
    heads, group = {}, None
    for c in range(1, g.max_col + 1):
        top, sub = cell(g, 2, c), cell(g, 3, c)
        if isinstance(top, str) and top.strip() and not top.startswith("From Date"):
            group = top.strip()
        if isinstance(sub, str) and sub.strip() and not sub.startswith("To Date"):
            heads[c] = f"{group} | {sub.strip()}" if group else sub.strip()
    return heads


def is_exec_layout(g: SheetGrid) -> bool:
    v = cell(g, 2, 3)
    return isinstance(v, str) and v.startswith("From Date")


# --------------------------------------------------------------------------------------
def profile_sheet(g: SheetGrid):
    exec_like = is_exec_layout(g)
    hdr = detect_header_rows(g) if not exec_like else {"period_rows": [2, 3], "working_day_row": None,
                                                         "caption_rows": [2, 3]}
    heads = exec_layout_headers(g) if exec_like else column_headers(g, hdr)
    first_data_row = (max(hdr["period_rows"] + hdr["caption_rows"] + [hdr["working_day_row"] or 0]) + 1
                      if not exec_like else 4)

    # monthly columns = a single YYYYMM in every period row (from == to)
    monthly_cols = []
    if not exec_like:
        for c in range(1, g.max_col + 1):
            pv = [cell(g, r, c) for r in hdr["period_rows"]]
            pv = [int(v) for v in pv if is_period(v)]
            if pv and len(set(pv)) == 1 and not any(isinstance(cell(g, r, c), str) and
                                                    re.search(r"(Act|Tar|Var|Diff)", cell(g, r, c) or "")
                                                    for r in hdr["caption_rows"] + [5]):
                monthly_cols.append((c, pv[0]))
    month_of = dict(monthly_cols)

    # ------------------------------ rows ------------------------------------------
    row_meta = []
    section, block, brand_ctx, blank_run = None, 0, None, 0
    for r in range(1, g.max_row + 1):
        vals = g.rows[r - 1]
        nonblank = [(i + 1, v) for i, v in enumerate(vals) if v not in (None, "")]
        if not nonblank:
            blank_run += 1
            continue
        gap_before, blank_run = blank_run, 0
        label_cols = [(c, v) for c, v in nonblank if isinstance(v, str) and c <= 4]
        nums = [(c, float(v)) for c, v in nonblank if is_num(v) and c > 4]
        label = label_cols[-1][1].strip() if label_cols else None
        tag = None
        if exec_like:
            b = cell(g, r, 2)
            tag = b.strip() if isinstance(b, str) else None
        if isinstance(label, str) and label.startswith("From Date"):
            block += 1
            brand_ctx = None
        # brand context
        nb = norm_brand(label) or norm_brand(tag)
        is_header_like = r < first_data_row
        if label and not nums and not is_header_like:
            section = label
            if nb:
                brand_ctx = nb
            elif gap_before >= 2 and not exec_like:
                brand_ctx = None          # a new major section starts -> forget the brand
        if label and label.endswith("Brand Total"):
            brand_ctx = norm_brand(label)
        if exec_like and label and label.startswith("Direct") and block in (1, 2, 3):
            # top (all-brand) blocks
            pass
        row_brand = nb or brand_ctx
        if month_of:                     # time series: stats on monthly columns only
            mnums = [(c, v) for c, v in nums if c in month_of]
        else:
            mnums = nums
        numvals = [v for _, v in mnums]
        first = numvals[0] if numvals else None
        last = numvals[-1] if numvals else None
        prior = numvals[-2] if len(numvals) > 1 else None
        mtype = metric_type(label if not exec_like else None) if not exec_like else "mixed (see columns)"
        flags = []
        if numvals:
            if any(v < 0 for v in numvals) and mtype in ("volume / count", "per-day volume"):
                flags.append("negative volume")
            if mtype == "rate / ratio" and any(v > 1.0 for v in numvals) and "distribution" not in (label or "").lower():
                flags.append("rate > 100%")
        if nums and not label:
            flags.append("numbers without a row label")
        row_meta.append({
            "Sheet": g.name, "Row": r, "Block #": block if exec_like else None,
            "Section": section, "Brand (inferred)": row_brand, "Tag (col B)": tag,
            "Item / Metric label": label, "Metric type": mtype,
            "Is header row": is_header_like,
            "Non-blank cells": len(nonblank), "Numeric cells": len(nums),
            "Text cells": sum(isinstance(v, str) for _, v in nonblank),
            "First numeric col": col_letter(nums[0][0]) if nums else None,
            "Last numeric col": col_letter(nums[-1][0]) if nums else None,
            "Stats basis": ("monthly columns" if month_of else "all numeric columns"),
            "First month": month_of.get(mnums[0][0]) if (month_of and mnums) else None,
            "Last month": month_of.get(mnums[-1][0]) if (month_of and mnums) else None,
            "Months with data": len(mnums) if month_of else None,
            "Min": min(numvals) if numvals else None,
            "Max": max(numvals) if numvals else None,
            "Mean": st.fmean(numvals) if numvals else None,
            "Sum": sum(numvals) if numvals else None,
            "First value": first, "Prior value": prior, "Last value": last,
            "Last vs prior %": ((last / prior - 1) if (prior not in (None, 0) and last is not None) else None),
            "Error-value cells": sum(1 for _, v in nonblank if is_err(v)),
            "Zero cells": sum(1 for v in numvals if v == 0),
            "Negative cells": sum(1 for v in numvals if v < 0),
            "Flags": "; ".join(flags) or None,
        })

    # ------------------------------ columns ---------------------------------------
    col_meta = []
    for c in range(1, g.max_col + 1):
        vals = [cell(g, r, c) for r in range(first_data_row, g.max_row + 1)]
        nonblank = [v for v in vals if v not in (None, "")]
        if not nonblank and c not in heads:
            continue
        nums = [float(v) for v in nonblank if is_num(v)]
        errs = [v for v in nonblank if is_err(v)]
        texts = [v for v in nonblank if isinstance(v, str) and not is_err(v)]
        dates = [v for v in nonblank if hasattr(v, "year")]
        period_vals = [cell(g, r, c) for r in hdr["period_rows"]]
        pv = [str(int(v)) for v in period_vals if is_period(v)]
        rate_share = (sum(1 for v in nums if -1.5 <= v <= 1.5 and not float(v).is_integer()) / len(nums)) if nums else None
        wd = cell(g, hdr["working_day_row"], c) if hdr["working_day_row"] else None
        col_meta.append({
            "Sheet": g.name, "Column": col_letter(c), "Col #": c,
            "Header (inferred)": heads.get(c),
            "Period from": pv[0] if pv else None, "Period to": pv[-1] if pv else None,
            "Working days (hdr)": wd if is_num(wd) else None,
            "Non-blank": len(nonblank), "Blank": len(vals) - len(nonblank),
            "Numeric": len(nums), "Text": len(texts), "Dates": len(dates), "Error values": len(errs),
            "Distinct values": len(set(map(str, nonblank))),
            "Min": min(nums) if nums else None, "Max": max(nums) if nums else None,
            "Mean": st.fmean(nums) if nums else None,
            "Median": st.median(nums) if nums else None,
            "Std dev": st.pstdev(nums) if len(nums) > 1 else None,
            "Zeros": sum(1 for v in nums if v == 0),
            "Negatives": sum(1 for v in nums if v < 0),
            "Share that look like rates": rate_share,
            "Sample text": "; ".join(list(dict.fromkeys(texts))[:3]) or None,
        })

    # ------------------------------ sheet -----------------------------------------
    all_vals = [v for row in g.rows for v in row if v not in (None, "")]
    types = Counter("number" if is_num(v) else "error" if is_err(v) else "text" if isinstance(v, str)
                    else "date" if hasattr(v, "year") else type(v).__name__ for v in all_vals)
    periods = sorted({str(int(v)) for r in hdr["period_rows"] for v in g.rows[r - 1] if is_period(v)})
    sections = [m["Section"] for m in row_meta if m["Section"]]
    brands = sorted({m["Brand (inferred)"] for m in row_meta if m["Brand (inferred)"]})
    sheet_meta = {
        "Sheet": g.name,
        "Layout": "Exec snapshot (metric groups x Actual/Projecting/Target/Var)" if exec_like
        else "Time series (rows = metrics, columns = periods)",
        "Used range": f"A1:{col_letter(g.max_col)}{g.max_row}",
        "Rows (used)": g.max_row, "Columns (used)": g.max_col,
        "Non-empty rows": len(row_meta),
        "Non-empty cells": len(all_vals),
        "Numeric cells": types.get("number", 0), "Text cells": types.get("text", 0),
        "Date cells": types.get("date", 0),
        "Error-value cells (#DIV/0! etc.)": types.get("error", 0),
        "Formula cells": g.formulas,
        "Merged ranges": g.merged,
        "Header / period rows": ", ".join(map(str, hdr["period_rows"])) or None,
        "Working-day row": hdr["working_day_row"],
        "First data row": first_data_row,
        "Periods covered": f"{periods[0]} - {periods[-1]} ({len(periods)} distinct)" if periods else None,
        "Monthly columns": (f"{col_letter(monthly_cols[0][0])}:{col_letter(monthly_cols[-1][0])} "
                            f"({monthly_cols[0][1]}-{monthly_cols[-1][1]}, n={len(monthly_cols)})"
                            if monthly_cols else None),
        "Distinct sections": len(set(sections)),
        "Brands found": ", ".join(brands) or None,
        "Report period (B2/B3)": (f"{cell(g, 2, 2)} - {cell(g, 3, 2)}" if exec_like else None),
        "Notes": "; ".join(x for x in [
            "All values are hard-coded (no live formulas) - checks must be re-derived" if g.formulas == 0 else None,
            f"{types.get('error', 0)} Excel error values pasted as values" if types.get("error") else None] if x) or None,
    }
    return sheet_meta, col_meta, row_meta


def profile_workbook(path: str):
    grids = load_grids(path)
    sheets, cols, rows = [], [], []
    for g in grids:
        s, c, r = profile_sheet(g)
        sheets.append(s); cols.extend(c); rows.extend(r)
    return sheets, cols, rows, grids


def write_tables(path_out: str, tables: dict[str, list[dict]]):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook(); wb.remove(wb.active)
    for name, data in tables.items():
        ws = wb.create_sheet(name)
        if not data:
            continue
        hdr = list(data[0].keys())
        ws.append(hdr)
        for d in data:
            ws.append([d.get(h) for h in hdr])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF", name="Arial")
            c.fill = PatternFill("solid", fgColor="1F4E78")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    wb.save(path_out)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else src.rsplit(".", 1)[0] + "_metadata.xlsx"
    s, c, r, _ = profile_workbook(src)
    write_tables(out, {"Meta_Sheets": s, "Meta_Columns": c, "Meta_Rows": r})
    print(f"Wrote {out}: {len(s)} sheets, {len(c)} columns, {len(r)} rows profiled")
