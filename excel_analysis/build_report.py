#!/usr/bin/env python3
"""
build_report.py - one command that produces the full deliverable.

    python build_report.py  UPdate.xlsx  UPdate_Analysis.xlsx

Output = a copy of the original workbook (all original sheets untouched) plus:
    READ ME                   guide, legend, how to re-run
    Key Findings              headline calculation errors + clear brand shifts (auto-ranked)
    Weighted Conv & Closing   how Conversion / Closing, their Targets and W Targets are built,
                              with LIVE formulas that re-derive the Exec - F27 July numbers,
                              a mix-vs-rate decomposition and an editable calculator
    Calc Errors               every failed check (cell, reported, re-derived, explanation)
    Brand x Item              brand x item x metric: July vs June vs F26 YTD vs targets
    Trend Shifts              last month vs prior / trailing-12 on every time-series row
    Meta - Sheets / Meta - Columns / Meta - Rows   metadata from profile_workbook.py
"""
from __future__ import annotations

import os
import sys
from datetime import date

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import analyse_update as AU          # noqa: E402
import profile_workbook as PW        # noqa: E402

ARIAL = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
RED_FILL = PatternFill("solid", fgColor="F8CBAD")
AMB_FILL = PatternFill("solid", fgColor="FFE699")
GRN_FILL = PatternFill("solid", fgColor="C6EFCE")
GREY_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_BLUE, F_GREEN, F_BLACK = "0000FF", "008000", "000000"
SRC = "'Exec - F27 July'"
PCT, PCT2, INT = "0.0%;-0.0%;-", "0.00%;-0.00%;-", "#,##0;(#,##0);-"


def font(bold=False, color=F_BLACK, size=10, italic=False):
    return Font(name=ARIAL, bold=bold, color=color, size=size, italic=italic)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w


def write_table(ws, data, start_row=1, widths=None, pct_cols=(), int_cols=(), wrap_cols=()):
    if not data:
        ws.cell(start_row, 1, "No rows").font = font(italic=True)
        return
    hdr = list(data[0].keys())
    for j, h in enumerate(hdr, 1):
        c = ws.cell(start_row, j, h)
        c.font = font(True, "FFFFFF"); c.fill = HDR_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, d in enumerate(data, start_row + 1):
        for j, h in enumerate(hdr, 1):
            v = d.get(h)
            if isinstance(v, float) and v != v:
                v = None
            c = ws.cell(i, j, v)
            if isinstance(v, str) and (v.startswith("=") or PW.is_err(v)):
                c.value = v; c.data_type = "s"      # keep as literal text, never a formula/error
            c.font = font()
            if h in pct_cols:
                c.number_format = PCT
            elif h in int_cols:
                c.number_format = INT
            if h in wrap_cols:
                c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = ws.cell(start_row + 1, 1)
    ws.auto_filter.ref = f"A{start_row}:{L(len(hdr))}{start_row + len(data)}"
    if widths:
        set_widths(ws, widths)
    ws.row_dimensions[start_row].height = 30


# =====================================================================================
def sheet_readme(wb, res, meta_sheets):
    ws = wb.create_sheet("READ ME", 0)
    ws.sheet_view.showGridLines = False
    f = res["findings"]
    hi = sum(1 for x in f if x["Severity"] == "High")
    md = sum(1 for x in f if x["Severity"] == "Medium")
    shifts = sum(1 for x in res["brand"] if x["Clear shift?"])
    tshifts = sum(1 for x in res["trends"] if x["Clear shift?"])
    lines = [
        ("UPdate workbook - metadata, data-quality review and weighted-target build-up", "title"),
        (f"Generated {date.today():%d %b %Y} by excel_analysis/build_report.py (re-run any time on a refreshed file).", "it"),
        ("", None),
        ("What was added (the 8 original sheets are unchanged)", "h"),
        ("Key Findings - headline calculation errors and the clearest brand shifts, ranked.", None),
        ("Weighted Conv & Closing - how Conversion %, Gross Closing %, Targets and W Targets are calculated, with live formulas "
         "that re-derive the Exec - F27 July numbers, a mix-vs-rate split and an editable calculator.", None),
        (f"Calc Errors - {len(f)} failed checks ({hi} High, {md} Medium). Every row gives the cell, the reported value, the "
         "re-derived value and what it means. Use the filters.", None),
        (f"Brand x Item - every brand x item x metric line: July vs June vs F26 YTD vs target. {shifts} lines flagged as a "
         "clear shift (sort by Shift score).", None),
        (f"Trend Shifts - every labelled row in the 4 trend sheets: last month vs prior month and vs trailing-12 mean. "
         f"{tshifts} flagged.", None),
        ("Meta - Sheets / Meta - Columns / Meta - Rows - metadata for every sheet, every used column and every non-empty row.", None),
        ("", None),
        ("Important facts about the source file", "h"),
        ("1. The workbook contains no live formulas (0 formula cells in all 8 sheets) - every %, variance and total is a "
         "hard-typed value, so nothing re-calculates. All checks here re-derive the numbers from the volumes.", None),
        ("1b. 2,317 cells hold Excel error codes pasted as values (#DIV/0!, #VALUE!, #REF!) in Month on Month, Summary and "
         "Hippo Branded. They are part of the original data and are listed in Calc Errors (category 'Error values').", None),
        ("2. Exec sheets: Projecting = Actual x 1.04545 (= 23 / 22 working days). Trend sheets use the July PROJECTION in the "
         "'202607' / '202607 Act' columns, not the actual.", None),
        ("3. Conversion % is measured only on conversion-eligible lines (where a contact precedes the lead, e.g. OQ-PNV, "
         "Hippo-Branded); OQ-PV, End-to-End, Hippo-Unbranded, Broker etc. carry contacts = leads and no conversion.", None),
        ("", None),
        ("Colour legend (new sheets)", "h"),
        ("Blue text = hard-coded input copied from the source / editable    Green text = live link to another sheet    "
         "Black text = formula    Yellow fill = cells you can change in the calculator", None),
        ("Severity: High = wrong number / broken logic that changes a reported KPI;  Medium = presentation, label or "
         "partial-period issue;  Low = rounding-sized or immaterial;  Info = expected structural gap (explained).", None),
        ("", None),
        ("How to re-run", "h"),
        ("python excel_analysis/build_report.py <input.xlsx> <output.xlsx>      (metadata only: python excel_analysis/profile_workbook.py <input.xlsx>)", None),
    ]
    r = 1
    for text, kind in lines:
        c = ws.cell(r, 1, text)
        if kind == "title":
            c.font = font(True, "1F4E78", 14)
        elif kind == "h":
            c.font = font(True, "FFFFFF", 11); c.fill = HDR_FILL
        elif kind == "it":
            c.font = font(italic=True, color="595959")
        else:
            c.font = font()
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        if kind is None and text and len(text) > 120:
            ws.row_dimensions[r].height = 15 * (1 + len(text) // 140)
        r += 1
    r += 1
    ws.cell(r, 1, "Sheet inventory (from Meta - Sheets)").font = font(True, "FFFFFF", 11)
    ws.cell(r, 1).fill = HDR_FILL
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    r += 1
    cols = ["Sheet", "Layout", "Used range", "Non-empty rows", "Numeric cells", "Formula cells", "Periods covered", "Brands found"]
    for j, h in enumerate(cols, 1):
        c = ws.cell(r, j, h); c.font = font(True); c.fill = SUB_FILL; c.border = BOX
    for s in meta_sheets:
        r += 1
        for j, h in enumerate(cols, 1):
            c = ws.cell(r, j, s.get(h)); c.font = font(); c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical="top")
    set_widths(ws, [24, 34, 14, 12, 12, 11, 26, 40])


# =====================================================================================
def find(res, **kw):
    for f in res["findings"]:
        if all((kw[k] in str(f.get(k, ""))) for k in kw):
            return f
    return None


def sheet_key_findings(wb, res):
    ws = wb.create_sheet("Key Findings", 1)
    ws.sheet_view.showGridLines = False
    set_widths(ws, [5, 30, 22, 16, 16, 16, 70])
    ws.cell(1, 1, "Key findings - calculation errors and clear shifts").font = font(True, "1F4E78", 14)
    ws.cell(2, 1, "Values pulled from the analysis at build time; IDs link to the 'Calc Errors' sheet.").font = font(italic=True, color="595959")

    curated = [
        ("Copy / link error", dict(Category="Copy / link error", Sheet="Exec - F27 July"),
         "Non-Direct contacts in the brand-split block show the Direct (Excl Outbound) figure (same error in June and YTD sheets). "
         "Brand rows underneath add to 63,801, the real Non-Direct is 69,901."),
        ("Weighted target collapse", dict(Category="Weighted target", Sheet="Exec - F27 July", Item="Non-Direct"),
         "Non-Direct W Target conversion is 0.40% vs a 12.8% plan, so 'W Tar Var' shows +3,220%. Cause: OEM (W 0.006%), 1Life (W 0%), "
         "Dealer Channel (W 0%) and Affinities Other (no target) feed 0% into the weighted average."),
        ("Weighted target collapse", dict(Category="Weighted target", Sheet="Exec - F27 July", Item="Affinities", Check="Gross Closing"),
         "Affinities W Target closing 16.1% vs 35.1% plan -> W Tar Var +115%. Lines with leads but a 0 W target (OEM, Affinities Other) pull it down."),
        ("Hierarchy", dict(Category="Hierarchy sum", Sheet="Exec - F27 July", Item="Tied Agency"),
         "Tied Agency line shows 0 while its three sub-lines carry 488 contacts / 112 gross sales. Non-Direct includes them, "
         "so Non-Direct gross sales (7,569) exceed the sum of the lines shown by 111."),
        ("Hierarchy", dict(Category="Hierarchy sum", Sheet="Exec - F27 July", Item="Total Personal Lines Incl MBO"),
         "Total contacts (238,632) are 9,932 more than Direct + Non-Direct (228,700). No line on the sheet explains the gap "
         "(June: 7,164)."),
        ("Brand reconciliation", dict(Category="Brand reconciliation", Sheet="Exec - F27 July", Item="Affinities Other"),
         "Affinities Other: 4,183 contacts on the all-brand page but only 140 across the five brand pages (Mortgage Market volume "
         "not allocated to a brand)."),
        ("Brand split", dict(Category="Brand split sum", Sheet="Exec - F27 July", Item="Direct (Excl Outbound)"),
         "Direct (Excl Outbound) = 81,533 contacts but the five brand rows add to 81,203 (330 not allocated to a brand)."),
        ("Rate > 100%", dict(Category="Sanity", Sheet="Exec - F27 July", Cell="L60"),
         "Virseker Non-Direct conversion 379% and gross effectiveness 101%. Contacts fell 82% (Insurance Exchange AFRI/SOLID), "
         "leads did not. The numerator and denominator come from different populations."),
        ("Period / label", dict(Category="Period / label", Sheet="Month on Month", Item="F26 YTD June Act"),
         "'F26 YTD June' columns in Month on Month and NEW - Month on Month cover Jul-25 to Jul-26 (13 months, including projected "
         "July): contacts 2,985,725 vs true F26 YTD 2,736,250; target 3,061,686 vs 2,839,930."),
        ("Period / label", dict(Category="Period / label", Check="PROJECTION"),
         "'202607 Act' in the trend sheets is the July projection (249,475), not the booked actual (238,632)."),
        ("Period / label", dict(Category="Period / label", Item="July 2026 working days"),
         "July has 23 working days in the trend sheets but 22 in Summary, so per-day rates differ by ~4.5% between sheets."),
        ("Error values", dict(Category="Error values", Sheet="Month on Month", Check="#REF!"),
         "Excel error codes pasted as values: Month on Month 771 cells (627 #DIV/0!, 141 #VALUE!, 3 #REF!), Summary 1,529 "
         "#DIV/0!, Hippo Branded 17. The #REF! cells (CT1120, CV1120, CR1172) mean the original number is lost."),
        ("Stray value", dict(Category="Stray value", Sheet="Exec - F27 July"),
         "Orphan numbers with no label (AF45 = 2,839; AF189 = -7,407) are left in the Exec sheets (all three months)."),
        ("Projection", dict(Category="Projection", Sheet="Exec - F27 July"),
         "Non-Direct 'Projecting' contacts (73,075) are inconsistent with the Actual on the same row (81,533 x 1.045 = 85,238). "
         "The Actual is the wrong link (see first item)."),
        ("Completeness", dict(Category="Completeness", Sheet="Exec - F27 July"),
         "'Total Personal Lines Excl MBO' shows conversion/closing % with no contacts or leads, so it cannot be verified."),
    ]
    r = 4
    hdr = ["#", "Topic", "Where", "Reported", "Re-derived", "Finding ID", "What it means"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(r, j, h); c.font = font(True, "FFFFFF"); c.fill = HDR_FILL
    n = 0
    for topic, crit, text in curated:
        kw = {k: v for k, v in crit.items() if k != "Cell"}
        f = None
        for x in res["findings"]:
            if all(v in str(x.get(k, "")) for k, v in kw.items()) and ("Cell" not in crit or crit["Cell"] in x["Cell / Row"]):
                f = x; break
        if not f:
            continue
        n += 1; r += 1
        vals = [n, topic, f"{f['Sheet']} {f['Cell / Row'][:40]}", f["Reported"], f["Expected / Re-derived"], f["ID"], text]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v); c.font = font(); c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if j in (4, 5) and isinstance(v, float) and abs(v) < 20:
                c.number_format = PCT2
            elif j in (4, 5) and isinstance(v, (int, float)):
                c.number_format = INT
        ws.row_dimensions[r].height = 45
    # ------------------ top shifts per brand -------------------------------------
    r += 2
    ws.cell(r, 1, "Clearest brand shifts - Exec July vs June vs F26 YTD (top 6 per brand by shift score, base >= 500)").font = font(True, "1F4E78", 12)
    r += 1
    ws.cell(r, 1, "Volumes compare July PROJECTED vs June actual; rates compare July actual vs June actual (pp). "
                  "Lines where the rate is > 100% are listed in Calc Errors and excluded here.").font = font(italic=True, color="595959")
    r += 1
    hdr2 = ["#", "Brand", "Item / Channel", "Metric", "July", "June", "Shift detail"]
    for j, h in enumerate(hdr2, 1):
        c = ws.cell(r, j, h); c.font = font(True, "FFFFFF"); c.fill = HDR_FILL
    def clean(b):
        if not b["Clear shift?"] or "data issue" in (b["Shift detail"] or ""):
            return False
        if b["Type"] == "rate" and any(isinstance(x, (int, float)) and x > 1 for x in
                                       (b["Jul Actual"], b["Jun Actual"], b["F26 YTD Actual"])):
            return False                       # rate > 100% somewhere = data issue, listed in Calc Errors
        if any(isinstance(x, (int, float)) and abs(x) > 3 for x in
               (b["Jul Var vs Tar"], b["Jun Var vs Tar"], b["F26 YTD Var vs Tar"])):
            return False                       # target too small to be meaningful
        return True
    rows, seen = [], set()
    for b in res["brand"]:
        key = (b["Brand"], b["Metric"], b["Jul Actual"], b["Jun Actual"])
        if clean(b) and key not in seen:       # e.g. 'Hippo' and 'Hippo - Branded' carrying identical rates
            seen.add(key); rows.append(b)
    for brand in ["All brands"] + AU.BRANDS:
        sel = sorted([b for b in rows if b["Brand"] == brand], key=lambda b: -(b["Shift score"] or 0))[:6]
        for b in sel:
            r += 1
            jul = b["Jul Projecting"] if b["Type"] == "vol" else b["Jul Actual"]
            vals = [None, brand, b["Item / Channel"], b["Metric"], jul, b["Jun Actual"], b["Shift detail"]]
            for j, v in enumerate(vals, 1):
                c = ws.cell(r, j, v); c.font = font(); c.border = BOX
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if j in (5, 6):
                    c.number_format = PCT if b["Type"] == "rate" else INT
            ws.row_dimensions[r].height = 30
    return ws


# =====================================================================================
def sheet_weighted(wb):
    ws = wb.create_sheet("Weighted Conv & Closing", 2)
    ws.sheet_view.showGridLines = False
    set_widths(ws, [34, 9, 14, 12, 14, 12, 12, 16, 16, 12, 44])
    r = 1
    ws.cell(r, 1, "How Conversion %, Gross Closing %, their Targets and Weighted (W) Targets are built").font = font(True, "1F4E78", 14)
    r += 1
    ws.cell(r, 1, "Source: Exec - F27 July (period 202607). Green = live link to that sheet, black = formula. "
                  "Change the source sheet and everything below re-calculates.").font = font(italic=True, color="595959")
    # ---------------- definitions ----------------
    r += 2
    defs = [
        ("Measure", "Formula (per line i)", "Notes"),
        ("Conversion %", "Eligible Leads / Eligible Contacts",
         "Only conversion-eligible lines (contact precedes lead): Traditional, Competition, OQ-PNV, Direct Marketing, Affiliates, "
         "Hippo-Branded, DM Outbound. OQ-PV / End-to-End / Hippo-Unbranded / Broker have contacts = leads and are excluded."),
        ("Conversion Target %", "Target eligible Leads / Target Contacts", "Plan rate at the PLAN channel mix."),
        ("W Target Conversion %", "SUM( Actual eligible Contacts_i x Target Conv_i ) / SUM( Actual eligible Contacts_i )",
         "Plan rate re-weighted to the ACTUAL channel mix. Built bottom-up: a parent's W = contact-weighted average of its "
         "children's W; at the lowest grain W_i = Target_i. A line with no target contributes 0% (see example D)."),
        ("Gross Closing %", "Gross Sales / Leads", "All leads (all lines)."),
        ("Closing Target %", "Target Gross Sales / Target Leads", "Plan closing at the plan lead mix."),
        ("W Target Closing %", "SUM( Actual Leads_i x Target Closing_i ) / SUM( Actual Leads_i )", "Plan closing at the actual lead mix."),
        ("Tar Var (rates)", "Actual / Target - 1", "Relative variance (not percentage points)."),
        ("W Tar Var", "Actual / W Target - 1", "Performance vs plan rates once the mix is neutralised."),
        ("Mix effect", "W Target - Target", "Change in the expected rate caused purely by channel mix moving vs plan."),
        ("Rate effect", "Actual - W Target", "Genuine conversion / closing performance. Mix + Rate = Actual - Target."),
        ("Gross Effectiveness %", "Eligible Gross Sales / Eligible Contacts  (= Conversion x Closing on eligible lines)", ""),
        ("NB Net CXD %", "1 - Net Sales / Gross Sales", ""),
        ("Projecting", "Actual x Total working days / Elapsed working days", "July: x 1.04545 (23/22)."),
        ("Tar Dif / Tar Var (volumes)", "Projecting - Target ;  (Projecting - Target) / Target", ""),
    ]
    for i, row in enumerate(defs):
        for j, v in enumerate(row):
            col = [1, 3, 8][j]
            c = ws.cell(r, col, v)
            c.font = font(i == 0, "FFFFFF" if i == 0 else F_BLACK)
            if i == 0:
                c.fill = HDR_FILL
            c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=11)
        ws.row_dimensions[r].height = 15 if i == 0 else (45 if len(row[2]) > 110 else 30)
        r += 1

    # ---------------- generic worked-example builder ----------------
    def example(r, title, note, vol_col, rate_col, tgt_col, w_col, lines, parent_row, vol_name, kind):
        ws.cell(r, 1, title).font = font(True, "FFFFFF", 11)
        for j in range(1, 12):
            ws.cell(r, j).fill = HDR_FILL
        r += 1
        ws.cell(r, 1, note).font = font(italic=True, color="595959")
        ws.cell(r, 1).alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
        ws.row_dimensions[r].height = 30
        r += 1
        hdr = ["Line (child)", "Src row", f"Actual {vol_name}", f"Actual {kind}", f"Implied numerator",
               f"Target {kind}", f"W Target {kind}", f"{vol_name} x W Target", f"{vol_name} x Target", "Mix share", "Comment"]
        for j, h in enumerate(hdr, 1):
            c = ws.cell(r, j, h); c.font = font(True); c.fill = SUB_FILL; c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 30
        first = r + 1
        for name, src, comment in lines:
            r += 1
            ws.cell(r, 1, name); ws.cell(r, 2, src)
            ws.cell(r, 3, f"=N({SRC}!{vol_col}{src})")
            ws.cell(r, 4, f"=N({SRC}!{rate_col}{src})")
            ws.cell(r, 5, f"=C{r}*D{r}")
            ws.cell(r, 6, f"=N({SRC}!{tgt_col}{src})")
            ws.cell(r, 7, f"=N({SRC}!{w_col}{src})")
            ws.cell(r, 8, f"=C{r}*G{r}")
            ws.cell(r, 9, f"=C{r}*F{r}")
            ws.cell(r, 10, f"=IF(SUM($C${first}:$C${first + len(lines) - 1})=0,0,C{r}/SUM($C${first}:$C${first + len(lines) - 1}))")
            ws.cell(r, 11, comment)
            for j in range(1, 12):
                c = ws.cell(r, j); c.border = BOX
                c.font = font(color=F_GREEN if j in (3, 4, 6, 7) else F_BLACK)
                c.number_format = {3: INT, 4: PCT2, 5: INT, 6: PCT2, 7: PCT2, 8: INT, 9: INT, 10: PCT}.get(j, "General")
            ws.cell(r, 11).alignment = Alignment(wrap_text=True, vertical="top")
        last = r
        r += 1
        tot = r
        ws.cell(r, 1, "Total / re-derived").font = font(True)
        ws.cell(r, 3, f"=SUM(C{first}:C{last})")
        ws.cell(r, 4, f"=IF(C{r}=0,0,E{r}/C{r})")
        ws.cell(r, 5, f"=SUM(E{first}:E{last})")
        ws.cell(r, 6, f"=IF(C{r}=0,0,I{r}/C{r})")
        ws.cell(r, 7, f"=IF(C{r}=0,0,H{r}/C{r})")
        ws.cell(r, 8, f"=SUM(H{first}:H{last})")
        ws.cell(r, 9, f"=SUM(I{first}:I{last})")
        ws.cell(r, 10, f"=SUM(J{first}:J{last})")
        ws.cell(r, 11, "D = SUM(E)/SUM(C) re-derived actual;  F = children's plan targets at actual mix;  G = SUMPRODUCT(C,W)/SUM(C) = re-derived W Target")
        for j in range(1, 12):
            c = ws.cell(r, j); c.border = BOX; c.fill = GREY_FILL
            c.font = font(True) if j < 11 else font(italic=True)
            c.number_format = {3: INT, 4: PCT2, 5: INT, 6: PCT2, 7: PCT2, 8: INT, 9: INT, 10: PCT}.get(j, "General")
        ws.cell(r, 11).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 40
        r += 1
        rep = r
        ws.cell(r, 1, f"Reported on Exec sheet (row {parent_row})").font = font(True)
        ws.cell(r, 2, parent_row)
        ws.cell(r, 3, f"=N({SRC}!{vol_col}{parent_row})")
        ws.cell(r, 4, f"=N({SRC}!{rate_col}{parent_row})")
        ws.cell(r, 6, f"=N({SRC}!{tgt_col}{parent_row})")
        ws.cell(r, 7, f"=N({SRC}!{w_col}{parent_row})")
        for j in (1, 2, 3, 4, 6, 7):
            c = ws.cell(r, j); c.border = BOX
            c.font = font(True, F_GREEN if j > 2 else F_BLACK)
            c.number_format = {3: INT, 4: PCT2, 6: PCT2, 7: PCT2}.get(j, "General")
        r += 1
        ws.cell(r, 1, "Difference (re-derived - reported)").font = font(True)
        ws.cell(r, 4, f"=D{tot}-D{rep}")
        ws.cell(r, 7, f"=G{tot}-G{rep}")
        ws.cell(r, 8, f'=IF(ABS(G{tot}-G{rep})<0.0005,"W Target reconciles","W Target does NOT reconcile")')
        for j in (4, 7):
            ws.cell(r, j).number_format = "+0.00%;-0.00%;0.00%"; ws.cell(r, j).font = font(True)
        ws.cell(r, 8).font = font(True)
        r += 1
        ws.cell(r, 1, "Variance split (reported figures)").font = font(True)
        r += 1
        items = [("Actual - Target (total gap, pp)", f"=D{rep}-F{rep}"),
                 ("Mix effect = W Target - Target (pp)", f"=G{rep}-F{rep}"),
                 ("Rate effect = Actual - W Target (pp)", f"=D{rep}-G{rep}"),
                 ("Check: Mix + Rate - Total (should be 0)", f"=(G{rep}-F{rep})+(D{rep}-G{rep})-(D{rep}-F{rep})"),
                 ("Tar Var = Actual / Target - 1", f"=IF(F{rep}=0,0,D{rep}/F{rep}-1)"),
                 ("W Tar Var = Actual / W Target - 1", f"=IF(G{rep}=0,0,D{rep}/G{rep}-1)")]
        for lab, fml in items:
            ws.cell(r, 1, lab).font = font()
            c = ws.cell(r, 4, fml); c.font = font(True); c.number_format = "+0.00%;-0.00%;0.00%"; c.border = BOX
            r += 1
        return r + 1, first, last, tot, rep

    r += 1
    r, *_ = example(
        r, "A. Direct - Conversion W Target (contacts-weighted)",
        "Children are the conversion-eligible lines of Direct (detail block rows 127-172). OQ uses only OQ-PNV and Hippo only "
        "Hippo-Branded, because only those lines have a conversion step. Competition has no target, so it enters at 0%.",
        "F", "L", "M", "P",
        [("Traditional Combined", 142, ""), ("Competitions Combined", 144, "No conversion target, so W = 0 and it dilutes the parent W"),
         ("OQ - PNV", 148, "Only OQ sub-line with a conversion step"), ("Direct Marketing", 154, ""),
         ("Affiliates", 152, ""), ("Hippo - Branded", 158, "Only Hippo sub-line with a conversion step"),
         ("DM Multibrand Outbound", 162, "Huge volume at an 8% rate, which drives the mix effect")],
        127, "Contacts", "Conv %")
    r, *_ = example(
        r, "B. Direct - Gross Closing W Target (leads-weighted)",
        "Every lead counts for closing. DM Multibrand Outbound already contains Referral, so Referral is not added again.",
        "S", "Y", "Z", "AC",
        [("Traditional Combined", 142, ""), ("Competitions Combined", 144, "No closing target, enters at 0%"),
         ("OQ", 146, "OQ's own W is itself a weighted average of PNV / PV / E2E (see C)"),
         ("Direct Marketing", 154, ""), ("Affiliates", 152, ""), ("Hippo", 156, ""),
         ("Better Compare", 168, ""), ("Cover Compare", 170, ""),
         ("Direct Other Clean Up", 172, "Leads with no target, enter at 0%"),
         ("DM Multibrand Outbound", 162, "")],
        127, "Leads", "Closing %")
    r, *_ = example(
        r, "C. OQ - Gross Closing W Target, one level down",
        "Shows the recursion: OQ's W is the lead-weighted average of its three sub-lines. The small residual comes from a lower "
        "grain (product / brand) that is not in the file.",
        "S", "Y", "Z", "AC",
        [("OQ - PNV", 148, ""), ("OQ - PV", 149, "Contacts = leads; closing on quotes"), ("OQ - End to End", 150, "Closing ~100% by design")],
        146, "Leads", "Closing %")
    r, _, _, d_tot, d_rep = example(
        r, "D. Where it breaks - Non-Direct Conversion W Target",
        "Several children have W = 0 although a plan target exists (Dealer Channel 61.3% target, W 0%) and Affinities' own W is "
        "0.44% vs an 8.7% target (OEM and 1Life W ~0). The parent W collapses to ~0.4% and W Tar Var reads +3,220%: a "
        "calculation error, not performance.",
        "F", "L", "M", "P",
        [("Affinities", 31, "W 0.44% vs target 8.73%: OEM / 1Life W ~0"), ("Bank", 32, "No conversion measured"),
         ("Dealer Channel", 33, "W = 0 although target = 61.3%"), ("Broker", 34, "Contacts = leads, no conversion"),
         ("Property Partners", 36, ""), ("Other Cleanup", 39, "")],
        29, "Contacts", "Conv %")
    c = ws.cell(r - 1, 1, f'="Corrected view: using each child\'s plan target instead of its broken W gives "&TEXT(F{d_tot},"0.0%")'
                          f'&" (children\'s targets at actual mix) vs the reported W of "&TEXT(G{d_rep},"0.00%")'
                          f'&", so the true W Tar Var is about "&TEXT(IF(F{d_tot}=0,0,D{d_rep}/F{d_tot}-1),"+0%;-0%")'
                          f'&", not "&TEXT(IF(G{d_rep}=0,0,D{d_rep}/G{d_rep}-1),"+0%;-0%")&"."')
    c.font = font(True, "C00000")
    ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=11)
    r += 1
    r, *_ = example(
        r, "E. Brand view - Budget Direct (Excl Outbound) Conversion W Target",
        "Same mechanics inside one brand page (Budget block rows 376-411).",
        "F", "L", "M", "P",
        [("Traditional Combined", 379, ""), ("Competitions Combined", 384, ""), ("OQ - PNV", 391, ""),
         ("Direct Marketing", 395, ""), ("Affiliates", 397, ""), ("Hippo - Branded", 401, "")],
        376, "Contacts", "Conv %")

    # ---------------- calculator ----------------
    ws.cell(r, 1, "F. Calculator - type your own lines (yellow cells) to see Target vs W Target vs Actual").font = font(True, "FFFFFF", 11)
    for j in range(1, 12):
        ws.cell(r, j).fill = HDR_FILL
    r += 1
    ws.cell(r, 1, "Pre-filled with Direct (Excl Outbound) by brand, Exec - F27 July rows 78-82 (values copied 202607). Overwrite "
                  "volumes, rates or targets to test scenarios. Rows with blank volume are ignored.").font = font(italic=True, color="595959")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
    r += 1
    hdr = ["Line", "", "Actual volume", "Actual rate", "Implied numerator", "Target rate", "Target volume",
           "Volume x Target rate", "Target numerator", "Mix share actual", "Mix share plan"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(r, j, h); c.font = font(True); c.fill = SUB_FILL; c.border = BOX
        c.alignment = Alignment(wrap_text=True)
    ws.row_dimensions[r].height = 30
    sample = [("Budget", 29584, 0.383421, 0.480416, 22308), ("First For Women", 18728, 0.384559, 0.402406, 21201),
              ("Dialdirect", 5964, 0.319966, 0.381645, 4346), ("Auto & General", 23976, 0.309797, 0.366840, 26661),
              ("Virseker", 2951, 0.417980, 0.487312, 2267), ("(spare line)", None, None, None, None)]
    first = r + 1
    for name, vol, rate, tgt, tvol in sample:
        r += 1
        ws.cell(r, 1, name)
        for j, v in ((3, vol), (4, rate), (6, tgt), (7, tvol)):
            c = ws.cell(r, j, v); c.fill = INPUT_FILL; c.font = font(color=F_BLUE)
        ws.cell(r, 5, f"=N(C{r})*N(D{r})")
        ws.cell(r, 8, f"=N(C{r})*N(F{r})")
        ws.cell(r, 9, f"=N(G{r})*N(F{r})")
        last = first + len(sample) - 1
        ws.cell(r, 10, f"=IF(SUM($C${first}:$C${last})=0,0,N(C{r})/SUM($C${first}:$C${last}))")
        ws.cell(r, 11, f"=IF(SUM($G${first}:$G${last})=0,0,N(G{r})/SUM($G${first}:$G${last}))")
        for j in range(1, 12):
            c = ws.cell(r, j); c.border = BOX
            if j not in (3, 4, 6, 7):
                c.font = font()
            c.number_format = {3: INT, 4: PCT2, 5: INT, 6: PCT2, 7: INT, 8: INT, 9: INT, 10: PCT, 11: PCT}.get(j, "General")
    ws.cell(first, 3).comment = Comment("Source: Exec - F27 July F78:F82 / L / M / H (values as at 202607). Editable input.", "analysis")
    last = r
    r += 1
    outs = [("Actual rate (volume-weighted)", f"=IF(SUM(C{first}:C{last})=0,0,SUM(E{first}:E{last})/SUM(C{first}:C{last}))"),
            ("Target rate (plan mix) = SUM(target numerators) / SUM(target volumes)", f"=IF(SUM(G{first}:G{last})=0,0,SUM(I{first}:I{last})/SUM(G{first}:G{last}))"),
            ("W Target rate (actual mix) = SUMPRODUCT(volume, target rate) / SUM(volume)", f"=IF(SUM(C{first}:C{last})=0,0,SUM(H{first}:H{last})/SUM(C{first}:C{last}))"),
            ("Mix effect = W Target - Target", None), ("Rate effect = Actual - W Target", None),
            ("Tar Var = Actual / Target - 1", None), ("W Tar Var = Actual / W Target - 1", None)]
    base = r
    for i, (lab, fml) in enumerate(outs):
        ws.cell(r, 1, lab).font = font(True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        if fml is None:
            a, t, w = f"D{base}", f"D{base + 1}", f"D{base + 2}"
            fml = {3: f"={w}-{t}", 4: f"={a}-{w}", 5: f"=IF({t}=0,0,{a}/{t}-1)", 6: f"=IF({w}=0,0,{a}/{w}-1)"}[i]
        c = ws.cell(r, 4, fml); c.font = font(True); c.border = BOX
        c.number_format = PCT2 if i < 3 else "+0.00%;-0.00%;0.00%"
        r += 1
    ws.cell(r, 1, "Compare: Exec - F27 July P76 (W Target) =").font = font(italic=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(r, 4, f"=N({SRC}!P76)"); c.font = font(True, F_GREEN); c.number_format = PCT2
    ws.cell(r, 5, f'="Weighting at BRAND level gives "&TEXT(D{base + 2},"0.0%")&" vs the reported "&TEXT(D{r},"0.0%")'
                  f'&". The sheet weights at sub-channel level inside each brand, so the grain of the weighting changes the '
                  f'answer. Always build W from the lowest grain upward."').font = font(italic=True, color="595959")
    ws.cell(r, 5).alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=11)
    ws.row_dimensions[r].height = 45
    ws.freeze_panes = "A4"
    return ws


# =====================================================================================
def add_flag_formatting(ws, col_letter, nrows, value, fill):
    ws.conditional_formatting.add(f"A2:{L(ws.max_column)}{nrows + 1}",
                                  FormulaRule(formula=[f'${col_letter}2="{value}"'], fill=fill))


def main(src, out):
    print("profiling ...")
    meta_s, meta_c, meta_r, _ = PW.profile_workbook(src)
    print("analysing ...")
    res = AU.run(src)
    print("writing ...")
    wb = openpyxl.load_workbook(src)
    sheet_readme(wb, res, meta_s)
    sheet_key_findings(wb, res)
    sheet_weighted(wb)

    # Calc Errors
    ws = wb.create_sheet("Calc Errors", 3)
    cols = ["ID", "Severity", "Category", "Sheet", "Cell / Row", "Occurrences", "Item", "Brand", "Check",
            "Reported", "Expected / Re-derived", "Difference", "Explanation / Impact"]
    data = [{k: f.get(k) for k in cols} for f in res["findings"]]
    write_table(ws, data, widths=[8, 9, 18, 18, 18, 8, 34, 14, 52, 12, 12, 12, 80],
                wrap_cols=("Check", "Explanation / Impact", "Cell / Row"))
    for i in range(2, len(data) + 2):
        for j in (10, 11, 12):
            c = ws.cell(i, j)
            if isinstance(c.value, float) and abs(c.value) < 50:
                c.number_format = "0.0000"
            elif isinstance(c.value, (int, float)):
                c.number_format = INT
    for sev, fill in (("High", RED_FILL), ("Medium", AMB_FILL), ("Low", GREY_FILL), ("Info", SUB_FILL)):
        ws.conditional_formatting.add(f"B2:B{len(data) + 1}", CellIsRule(operator="equal", formula=[f'"{sev}"'], fill=fill))

    # Brand x Item
    ws = wb.create_sheet("Brand x Item", 4)
    bcols = list(res["brand"][0].keys())
    write_table(ws, res["brand"], widths=[15, 30, 20, 6] + [12] * 13 + [9, 9, 9, 60, 8],
                wrap_cols=("Shift detail",))
    rate_metric = {i + 2: (b["Type"] == "rate") for i, b in enumerate(res["brand"])}
    pct_hdr = {"Jul Var vs Tar", "Jun Var vs Tar", "F26 YTD Var vs Tar", "Jul vs Jun (% vol / pp rate)", "Jul vs F26 avg (% vol / pp rate)"}
    val_hdr = {"Jul Actual", "Jul Projecting", "Jul Target", "Jun Actual", "Jun Target", "F26 YTD Actual", "F26 YTD Target", "Base size (Jul)"}
    for j, h in enumerate(bcols, 1):
        for i in range(2, len(res["brand"]) + 2):
            c = ws.cell(i, j)
            if h in pct_hdr:
                c.number_format = PCT
            elif h in val_hdr:
                c.number_format = PCT if (rate_metric[i] and h != "Base size (Jul)") else INT
    sc = L(bcols.index("Clear shift?") + 1)
    add_flag_formatting(ws, sc, len(res["brand"]), "YES", AMB_FILL)

    # Trend shifts
    ws = wb.create_sheet("Trend Shifts", 5)
    write_table(ws, res["trends"], widths=[18, 6, 30, 15, 30, 7, 9, 12, 12, 12, 12, 12, 12, 9, 9, 8, 40],
                wrap_cols=())
    tcols = list(res["trends"][0].keys())
    for i, t in enumerate(res["trends"], 2):
        fmt = PCT if t["Type"] == "rate" else INT
        for h in ("Last value", "Prior value", "Trailing-12 mean", "Trailing-12 std"):
            ws.cell(i, tcols.index(h) + 1).number_format = fmt
        for h in ("Change vs prior (% vol / pp rate)", "Change vs T12 mean (% vol / pp rate)"):
            ws.cell(i, tcols.index(h) + 1).number_format = PCT
        for h in ("Z-score vs T12", "Step z (last 3 vs prior 12)"):
            ws.cell(i, tcols.index(h) + 1).number_format = "0.0"
    add_flag_formatting(ws, L(tcols.index("Clear shift?") + 1), len(res["trends"]), "YES", AMB_FILL)

    # Metadata
    for name, data, w in (("Meta - Sheets", meta_s, 22), ("Meta - Columns", meta_c, 14), ("Meta - Rows", meta_r, 14)):
        ws = wb.create_sheet(name)
        write_table(ws, data, widths=[w] * len(data[0]))
        if name == "Meta - Rows":
            ws.column_dimensions["G"].width = 34; ws.column_dimensions["D"].width = 30
        if name == "Meta - Sheets":
            for col in range(1, len(data[0]) + 1):
                ws.column_dimensions[L(col)].width = 24
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    c.alignment = Alignment(wrap_text=True, vertical="top")

    # tab colours
    for n, col in (("READ ME", "1F4E78"), ("Key Findings", "C00000"), ("Weighted Conv & Closing", "00B050"),
                   ("Calc Errors", "C00000"), ("Brand x Item", "FFC000"), ("Trend Shifts", "FFC000")):
        wb[n].sheet_properties.tabColor = col
    for n in ("Meta - Sheets", "Meta - Columns", "Meta - Rows"):
        wb[n].sheet_properties.tabColor = "7F7F7F"
    wb.active = 0
    wb.save(out)
    print(f"saved {out}")
    return res


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
