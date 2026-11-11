#!/usr/bin/env python3
"""
analyse_update.py - domain checks for the 'UPdate' channel-performance workbook.

What it does
------------
1. Re-derives every calculated column of the three Exec sheets (the file contains no
   live formulas) and lists each place where the reported number does not match:
       Tar Dif = Projecting - Target          Tar Var = Tar Dif / Target
       Rate Tar Var = Actual / Target - 1     W Tar Var = Actual / W Target - 1
       Gross Closing = Gross Sales / Leads    NB Net CXD = 1 - Net Sales / Gross Sales
       Projecting = Actual x (period factor)  parent volume = sum of child volumes
       sum of brands = all-brand line         duplicated lines agree with each other
   plus sanity rules (rates > 100%, negative sales / targets, weighted target collapsed
   to 0, stray numbers without labels, totals with rates but no volumes).
2. Re-derives the same identities on the month-on-month sheets (per-day = volume /
   working days, closing, CXD, effectiveness, target variances) and cross-checks the
   sheets against each other (June Exec vs Month-on-Month June, YTD columns, July).
3. Builds a brand x item x metric comparison across Exec July / June / June YTD and a
   trend-shift scan on every brand row in the time-series sheets, flagging clear shifts.

Usage:  python analyse_update.py UPdate.xlsx      (prints a summary)
The functions are imported by build_report.py.
"""
from __future__ import annotations

import math
import statistics as st
import sys
from collections import defaultdict

import openpyxl
from openpyxl.utils import column_index_from_string as ci, get_column_letter as L

EXEC_SHEETS = ["Exec - F27 July", "Exec - F26 June", "Exec - F26 June YTD"]
BRANDS = ["Budget", "First For Women", "Dialdirect", "Auto & General", "Virseker"]
BRAND_TAGS = {"budget": "Budget", "ffw": "First For Women", "first for women": "First For Women",
              "dial": "Dialdirect", "dialdirect": "Dialdirect", "a&g": "Auto & General",
              "auto & general": "Auto & General", "virseker": "Virseker",
              "tih multibrand": "TIH Multibrand"}

# Exec column map -------------------------------------------------------------------
VOL_GROUPS = {   # name: (Actual, Projecting, Target, Tar Dif, Tar Var)
    "Contacts": ("F", "G", "H", "I", "J"),
    "Leads": ("S", "T", "U", "V", "W"),
    "Gross Sales": ("AF", "AG", "AH", "AI", "AJ"),
    "NB Net Sales": ("AT", "AU", "AV", "AW", "AX"),
}
RATE_GROUPS = {  # name: (Actual, Target, Tar Var, W Target, W Tar Var)
    "Conversion %": ("L", "M", "N", "P", "Q"),
    "Gross Closing %": ("Y", "Z", "AA", "AC", "AD"),
    "Gross Effectiveness %": ("AL", "AM", "AN", None, None),
    "NB Net CXD %": ("AP", "AQ", "AR", None, None),
    "Net Effectiveness %": ("AZ", "BA", "BB", None, None),
}
EXEC_COLS = sorted({c for g in VOL_GROUPS.values() for c in g} |
                   {c for g in RATE_GROUPS.values() for c in g if c}, key=ci)

# Parent = sum(children) rules (children missing in a block count as 0; rule is applied
# only when at least one child is present in the same block after the parent).
HIERARCHY = [
    ("Total Personal Lines Incl MBO", ["Direct", "Non-Direct"]),
    ("Direct", ["Direct (Excl Outbound)", "DM Multibrand Outbound"]),
    ("Direct (Incl Outbound)", ["Direct (Excl Outbound)", "DM Multibrand Outbound"]),
    ("Direct (Excl Outbound)", ["Traditional Combined", "Competitions Combined", "OQ", "Direct Marketing",
                                "Affiliates", "Hippo", "Better Compare", "Cover Compare",
                                "Direct Other Clean Up"]),
    ("DM Multibrand Outbound", ["DM - Branded Outbound", "DM - Multibrand Outbound", "Referral"]),
    ("Traditional Combined", ["Generic - Traditional", "E-Commerce - Traditional"]),
    ("Competitions Combined", ["Generic - Competition", "E-Commerce - Competition"]),
    ("Generic", ["Generic - Traditional", "Generic - Competition"]),
    ("E-Commerce", ["E-Commerce - Traditional", "E-Commerce - Competition"]),
    ("OQ", ["OQ - PNV", "OQ - PV", "OQ - End to End"]),
    ("Hippo", ["Hippo - Branded", "Hippo - Unconverted", "Hippo - Unbranded"]),
    ("Non-Direct", ["Affinities", "Bank", "Dealer Channel", "Broker", "Sales Partnerships",
                    "Property Partners", "Tied Agency", "Other Cleanup"]),
    ("Dealer Channel", ["Multibrand", "Dealer Groups", "Independent Dealers", "Virtual Ops",
                        "Zambesi Dealers", "Provincial Groups", "Dealer Cleanup"]),
    ("Broker", ["Traditional Broker", "Dealer Broker", "OCC Broker", "Broker Bank"]),
    ("Property Partners", ["Balwin", "Landsdowne", "Waterfall"]),
    ("Tied Agency", ["Tied Agency Academy", "Tied Agency F2F", "Tied Agency Virtual",
                     "Agency F2F", "Agency Lite", "Agency Virtual"]),
    ("Bank", ["Nedbank Budget", "Nedbank FB", "Nedbank BB", "NGIB", "SBIB", "ABSA"]),
    ("Affinities", ["Insurance Exchange", "OEM", "Fidelity ADT", "1Life Motor and Home",
                    "Virgin Money MQ", "Affinities Other", "Kandua", "Namola"]),
    ("OEM", ["MyToyota Insurance", "Volkswagen Financial Services", "Audi Financial Services",
             "Toyota Financial Services"]),
]
SUM_COLS = {"Contacts": "F", "Leads": "S", "Gross Sales": "AF", "NB Net Sales": "AT",
            "Contacts Target": "H", "Leads Target": "U", "Gross Sales Target": "AH",
            "Net Sales Target": "AV"}


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, float) and math.isnan(v))


def F(sev, cat, sheet, ref, item, brand, check, reported, expected, note):
    diff = (reported - expected) if (num(reported) and num(expected)) else None
    return {"Severity": sev, "Category": cat, "Sheet": sheet, "Cell / Row": ref, "Item": item,
            "Brand": brand, "Check": check, "Reported": reported, "Expected / Re-derived": expected,
            "Difference": diff, "Explanation / Impact": note}


# =====================================================================================
#  EXEC SHEETS
# =====================================================================================
class ExecSheet:
    def __init__(self, ws_rows, name):
        self.name = name
        self.rows = {}
        self.order = []
        block, brand, parent = 0, None, None
        for i, r in enumerate(ws_rows, 1):
            lab = r[2] if len(r) > 2 else None
            tag = r[1] if len(r) > 1 else None
            vals = {c: (r[ci(c) - 1] if ci(c) - 1 < len(r) else None) for c in EXEC_COLS}
            has_num = any(num(v) for v in vals.values())
            if isinstance(lab, str) and lab.startswith("From Date"):
                block += 1; brand = None; parent = None
                continue
            if isinstance(lab, str) and lab.startswith("To Date"):
                continue
            if isinstance(lab, str) and lab.endswith("Brand Total"):
                brand = BRAND_TAGS.get(lab.replace(" - Brand Total", "").strip().lower())
            lab_s = lab.strip() if isinstance(lab, str) else None
            row_brand = BRAND_TAGS.get((tag or "").strip().lower()) if isinstance(tag, str) else None
            is_brand_row = row_brand is not None and lab_s is not None and \
                BRAND_TAGS.get(lab_s.lower()) == row_brand
            if lab_s and not is_brand_row:
                parent = lab_s
            if not lab_s and not has_num:
                continue
            self.rows[i] = dict(row=i, label=lab_s, tag=tag, block=block, block_brand=brand,
                                split_brand=row_brand if is_brand_row else None,
                                split_parent=parent if is_brand_row else None, v=vals)
            self.order.append(i)
        self.period = (ws_rows[1][1], ws_rows[2][1])

    def blocks(self):
        b = defaultdict(list)
        for i in self.order:
            b[self.rows[i]["block"]].append(self.rows[i])
        return b


def proj_factor(sh: ExecSheet):
    fs = sorted(r["v"]["G"] / r["v"]["F"] for r in sh.rows.values()
                if num(r["v"]["F"]) and num(r["v"]["G"]) and r["v"]["F"] > 2000)
    return st.median(fs) if fs else 1.0


def exec_checks(sh: ExecSheet):
    out = []
    S = sh.name
    fac = proj_factor(sh)
    for r in sh.rows.values():
        v, lab, row = r["v"], r["label"], r["row"]
        brand = r["block_brand"] or r["split_brand"]
        item = lab if not r["split_brand"] else f"{r['split_parent']} > {lab}"
        has_any = any(num(x) for x in v.values())
        if not lab and has_any:
            cells = ", ".join(f"{c}{row}={v[c]}" for c in EXEC_COLS if num(v[c]))
            out.append(F("Medium", "Stray value", S, f"Row {row}", "(no label)", brand,
                         "Number without a row label", cells, None,
                         "Orphan number sitting between blocks - looks like a leftover/hard-typed "
                         "value. It is not part of any line and should be deleted or labelled."))
            continue
        # ---- volume identities --------------------------------------------------
        for g, (a, p, t, d, var) in VOL_GROUPS.items():
            A, P, T, D, V = (v[a], v[p], v[t], v[d], v[var])
            if num(P) and num(T) and num(D) and abs((P - T) - D) > 1.01:
                out.append(F("High", "Arithmetic", S, f"{d}{row}", item, brand, f"{g}: Tar Dif = Projecting - Target",
                             D, P - T, "Target difference does not equal Projecting minus Target."))
            if num(D) and num(T) and T != 0 and num(V) and abs(D / T - V) > 0.0006:
                out.append(F("High", "Arithmetic", S, f"{var}{row}", item, brand, f"{g}: Tar Var = Tar Dif / Target",
                             V, D / T, "Target variance % does not equal Tar Dif / Target."))
            if num(A) and num(P) and A >= 20 and abs(P - A * fac) > max(3.01, 0.01 * A):
                out.append(F("High", "Projection", S, f"{p}{row}", item, brand,
                             f"{g}: Projecting = Actual x {fac:.5f}", P, round(A * fac),
                             "Projected value is inconsistent with the run-rate factor used on every other "
                             "line - the Actual (or the Projecting) on this line is pulled from a different row."))
            if num(T) and T < 0:
                out.append(F("Medium", "Sanity", S, f"{t}{row}", item, brand, f"{g}: target must be >= 0", T, 0,
                             "Negative target - a target cannot be negative; the % variance next to it is meaningless."))
            if num(A) and A < 0:
                out.append(F("Medium", "Sanity", S, f"{a}{row}", item, brand, f"{g}: actual must be >= 0", A, 0,
                             "Negative sales (cancellations exceed new business on a tiny base). "
                             "Ratios built on this line are distorted."))
        # ---- rate identities ----------------------------------------------------
        for g, (a, t, var, w, wvar) in RATE_GROUPS.items():
            A, T, V = v[a], v[t], v[var]
            if num(A) and num(T) and T != 0 and num(V) and abs(A / T - 1 - V) > 0.0006:
                out.append(F("High", "Arithmetic", S, f"{var}{row}", item, brand, f"{g}: Tar Var = Actual / Target - 1",
                             V, A / T - 1, "Rate variance does not equal Actual / Target - 1."))
            if w:
                W, WV = v[w], v[wvar]
                if num(A) and num(W) and W != 0 and num(WV) and abs(A / W - 1 - WV) > 0.0006:
                    out.append(F("High", "Arithmetic", S, f"{wvar}{row}", item, brand,
                                 f"{g}: W Tar Var = Actual / W Target - 1", WV, A / W - 1,
                                 "Weighted-target variance does not equal Actual / W Target - 1."))
                if num(W) and W == 0 and num(T) and T > 0:
                    out.append(F("High", "Weighted target", S, f"{w}{row}", item, brand,
                                 f"{g}: W Target collapsed to 0 while Target = {T:.1%}", W, T,
                                 "The weighted target is 0 although a plan target exists, so the W variance is "
                                 "blank/unusable and the line contributes 0 to every parent weighted target "
                                 "(understates the parent W Target)."))
                if num(W) and num(T) and T > 0 and W > 0 and (W / T < 0.5 or W / T > 2):
                    out.append(F("High", "Weighted target", S, f"{w}{row}", item, brand,
                                 f"{g}: W Target {W:.2%} vs Target {T:.2%} (x{W / T:.2f})", W, T,
                                 "Weighted target is implausibly far from the plan target - child lines with "
                                 "volume but a 0 / missing target are being averaged in at 0%. The resulting "
                                 f"W Tar Var ({(A / W - 1) if num(A) and W else float('nan'):+.0%}) is not a real "
                                 "performance signal."))
        # closing / CXD re-derivation
        if num(v["S"]) and v["S"] > 0 and num(v["AF"]) and num(v["Y"]) and abs(v["AF"] / v["S"] - v["Y"]) > 0.0006:
            out.append(F("High", "Arithmetic", S, f"Y{row}", item, brand, "Gross Closing = Gross Sales / Leads",
                         v["Y"], v["AF"] / v["S"], "Closing % does not equal Gross Sales / Leads."))
        if num(v["U"]) and v["U"] > 0 and num(v["AH"]) and num(v["Z"]) and abs(v["AH"] / v["U"] - v["Z"]) > 0.0006:
            out.append(F("High", "Arithmetic", S, f"Z{row}", item, brand, "Closing Target = Sales Target / Leads Target",
                         v["Z"], v["AH"] / v["U"], "Closing target is not Sales target / Leads target."))
        if num(v["AF"]) and v["AF"] > 0 and num(v["AT"]) and num(v["AP"]) and abs(1 - v["AT"] / v["AF"] - v["AP"]) > 0.0006:
            out.append(F("High", "Arithmetic", S, f"AP{row}", item, brand, "NB Net CXD = 1 - Net Sales / Gross Sales",
                         v["AP"], 1 - v["AT"] / v["AF"], "CXD % does not equal 1 - Net/Gross."))
        if num(v["AH"]) and v["AH"] > 0 and num(v["AV"]) and num(v["AQ"]) and abs(1 - v["AV"] / v["AH"] - v["AQ"]) > 0.0006:
            out.append(F("Medium", "Arithmetic", S, f"AQ{row}", item, brand, "CXD Target = 1 - Net Target / Gross Target",
                         v["AQ"], 1 - v["AV"] / v["AH"], "CXD target is not consistent with the net and gross sales targets."))
        # sanity on rates
        for g, col, lim in [("Conversion %", "L", 1.0), ("Gross Effectiveness %", "AL", 1.0),
                            ("Net Effectiveness %", "AZ", 1.0), ("Gross Closing %", "Y", 1.0)]:
            x = v[col]
            if num(x) and x > lim + 1e-9:
                out.append(F("High" if col in ("L", "AL", "AZ") else "Low", "Sanity", S, f"{col}{row}", item, brand,
                             f"{g} above 100%", x, lim,
                             "More leads than contacts / more sales than contacts or leads. Numerator and "
                             "denominator come from different populations (e.g. leads without a matching "
                             "contact, or contacts captured in another line) - the rate is not usable."))
        # rates without volumes
        if any(num(v[c]) for c in ("L", "Y")) and not num(v["F"]) and not num(v["S"]) and lab:
            out.append(F("Medium", "Completeness", S, f"Row {row}", item, brand,
                         "Rates shown but Contacts / Leads volumes are blank", None, None,
                         "Line shows conversion/closing % but no underlying contacts/leads, so the rate cannot "
                         "be verified or re-weighted."))
    return out


def exec_hierarchy_checks(sh: ExecSheet):
    out = []
    for b, rows in sh.blocks().items():
        bylab = {}
        for r in rows:
            if r["label"] and not r["split_brand"] and r["label"] not in bylab:
                bylab[r["label"]] = r
        for parent, kids in HIERARCHY:
            if parent not in bylab:
                continue
            p = bylab[parent]
            present = [bylab[k] for k in kids if k in bylab and bylab[k]["row"] > p["row"]]
            if parent in ("Total Personal Lines Incl MBO", "Direct", "Direct (Incl Outbound)", "Non-Direct") or \
                    parent.endswith("Brand Total"):
                present = [bylab[k] for k in kids if k in bylab]
            if len(present) < 2:
                continue
            for name, col in SUM_COLS.items():
                pv = p["v"][col]
                if not num(pv):
                    continue
                s = sum(k["v"][col] for k in present if num(k["v"][col]))
                if abs(pv - s) > 1.01:
                    missing = [k for k in kids if k not in bylab]
                    material = abs(pv - s) > max(5, 0.01 * abs(pv))
                    sev = ("High" if not missing else "Medium") if material else "Low"
                    if pv == 0 and s > 0:
                        sev = "High"
                    out.append(F(sev,
                                 "Hierarchy sum", sh.name, f"{col}{p['row']}", parent, p["block_brand"],
                                 f"{name}: {parent} = sum of {', '.join(k['label'] for k in present)}",
                                 pv, s,
                                 "Parent line does not equal the sum of its child lines in the same block"
                                 + (f" (children not shown in block: {', '.join(missing)})" if missing else "")
                                 + "."))
    return out


def exec_split_checks(sh: ExecSheet):
    """Brand-split groups: a parent line followed by Budget/FFW/Dial/A&G/Virseker(/TIH) rows."""
    out = []
    groups = defaultdict(list)
    for r in sh.rows.values():
        if r["split_brand"]:
            groups[(r["block"], r["split_parent"])].append(r)
    for (b, parent), kids in groups.items():
        prow = max((r for r in sh.rows.values() if r["block"] == b and r["label"] == parent
                    and not r["split_brand"] and r["row"] < kids[0]["row"]), key=lambda r: r["row"], default=None)
        if prow is None:
            continue
        for name, col in SUM_COLS.items():
            pv = prow["v"][col]
            if not num(pv):
                continue
            s = sum(k["v"][col] for k in kids if num(k["v"][col]))
            if abs(pv - s) > 1.01:
                out.append(F("High" if abs(pv - s) > 0.01 * abs(pv) else "Medium", "Brand split sum", sh.name, f"{col}{prow['row']}", parent, None,
                             f"{name}: {parent} = sum of brand rows {kids[0]['row']}-{kids[-1]['row']}",
                             pv, s, "The brand rows underneath do not add up to the line they split "
                             f"(difference {pv - s:+,.0f}) - volume not allocated to any brand, or a wrong link."))
    return out


def exec_duplicate_checks(sh: ExecSheet):
    """The same all-brand line is repeated in several blocks - all copies must agree."""
    out = []
    groups = defaultdict(list)
    for r in sh.rows.values():
        if r["label"] and not r["block_brand"]:
            key = (r["split_parent"], r["label"]) if r["split_brand"] else (None, r["label"])
            groups[key].append(r)
    for key, rs in groups.items():
        if len(rs) < 2:
            continue
        base = rs[0]
        for other in rs[1:]:
            for name, col in [("Contacts", "F"), ("Leads", "S"), ("Gross Sales", "AF"), ("NB Net Sales", "AT"),
                              ("Contacts Projecting", "G"), ("Conversion %", "L"), ("Gross Closing %", "Y")]:
                a, b = base["v"][col], other["v"][col]
                if num(a) and num(b) and abs(a - b) > (0.0006 if col in ("L", "Y") else 0.5):
                    # which copy is wrong? the one that breaks its own identities
                    out.append(F("High", "Copy / link error", sh.name, f"{col}{other['row']}",
                                 (f"{key[0]} > " if key[0] else "") + key[1], other["split_brand"],
                                 f"{name}: row {other['row']} should equal row {base['row']} (same line repeated)",
                                 b, a, "The same line appears in two places with different values - one cell is "
                                 "linked to the wrong source row (e.g. Non-Direct contacts showing the Direct "
                                 "(Excl Outbound) figure)."))
    return out


def exec_brand_recon(sh: ExecSheet):
    """Sum of the five brand blocks vs the all-brand line with the same label."""
    out = []
    allb, perb = {}, defaultdict(dict)
    for r in sh.rows.values():
        if not r["label"] or r["split_brand"]:
            continue
        if r["block_brand"]:
            lab = "Total Personal Lines Incl MBO" if r["label"].endswith("Brand Total") else r["label"]
            perb[lab].setdefault(r["block_brand"], r)
        else:
            allb.setdefault(r["label"], r)
    mb_lines = {"Total Personal Lines Incl MBO", "Direct", "DM Multibrand Outbound", "DM - Multibrand Outbound"}
    for lab, br in perb.items():
        if lab not in allb or len(br) < 5:
            continue
        a = allb[lab]
        for name, col in [("Contacts", "F"), ("Leads", "S"), ("Gross Sales", "AF"), ("NB Net Sales", "AT"),
                          ("Contacts Target", "H"), ("Gross Sales Target", "AH")]:
            if not num(a["v"][col]):
                continue
            s = sum(r["v"][col] for r in br.values() if num(r["v"][col]))
            d = a["v"][col] - s
            if abs(d) > 1.01:
                expl = ("Gap = Multibrand (TIH) volume that is not allocated to a brand - expected, but the "
                        "brand pages therefore do not add to the group total." if lab in mb_lines else
                        "Brand pages do not add up to the all-brand line - volume on the group line that is "
                        "not allocated to any brand page (or double counted).")
                out.append(F("Info" if lab in mb_lines else ("High" if abs(d) > max(5, 0.01 * abs(a["v"][col])) else "Low"),
                             "Brand reconciliation", sh.name, f"{col}{a['row']}", lab, "All brands",
                             f"{name}: all-brand row {a['row']} vs sum of 5 brand pages", a["v"][col], s, expl))
    return out


# =====================================================================================
#  TIME-SERIES SHEETS
# =====================================================================================
TS_SHEETS = {"Month on Month": dict(days_row=2, from_row=3, to_row=4, cap_row=5, label_col=3),
             "NEW - Month on Month": dict(days_row=1, from_row=2, to_row=3, cap_row=4, label_col=3),
             "Hippo Branded": dict(days_row=2, from_row=3, to_row=4, cap_row=5, label_col=3),
             "Online Quoting": dict(days_row=2, from_row=3, to_row=4, cap_row=5, label_col=2)}


class TSSheet:
    def __init__(self, rows, name, cfg):
        self.name, self.cfg, self.grid = name, cfg, rows
        g = lambda r, c: rows[r - 1][c - 1] if r - 1 < len(rows) and c - 1 < len(rows[r - 1]) else None
        self.g = g
        mc = max(len(r) for r in rows[:6])
        self.month_cols, self.special = [], {}
        for c in range(1, mc + 1):
            f, t, cap = g(cfg["from_row"], c), g(cfg["to_row"], c), g(cfg["cap_row"], c)
            if num(f) and num(t) and f == t and not (isinstance(cap, str) and any(k in cap for k in ("Act", "Tar", "Var", "Diff"))):
                self.month_cols.append((c, int(f)))
            if isinstance(cap, str) and cap.strip():
                self.special[cap.strip()] = c
        self.days = {c: g(cfg["days_row"], c) for c in range(1, mc + 1)}
        # rows: section / brand context
        self.items = []
        section, brand, blank = None, None, 0
        lc = cfg["label_col"]
        for i, r in enumerate(rows, 1):
            if i <= cfg["cap_row"]:
                continue
            vals = [x for x in r if x not in (None, "")]
            if not vals:
                blank += 1
                continue
            lab = r[lc - 1] if lc - 1 < len(r) else None
            tag = r[1] if len(r) > 1 else None
            lab = lab.strip() if isinstance(lab, str) else None
            has_num = any(num(x) for x in r[lc:])
            b = BRAND_TAGS.get((lab or "").lower()) or (BRAND_TAGS.get(tag.strip().lower()) if isinstance(tag, str) else None)
            if lab and not has_num:
                if b:
                    brand = b
                elif blank >= 2:
                    brand = None
                section = lab
            elif lab and has_num and b and lab.lower() in BRAND_TAGS:
                pass
            if lab and has_num and isinstance(tag, str) and BRAND_TAGS.get(tag.strip().lower()) and lc == 3:
                brand = BRAND_TAGS.get(tag.strip().lower())
                section = lab
            blank = 0
            if lab and has_num:
                self.items.append(dict(row=i, label=lab, section=section,
                                       brand=(BRAND_TAGS.get(lab.lower()) or brand), r=r))

    def series(self, it):
        return [(c, p, it["r"][c - 1]) for c, p in self.month_cols if c - 1 < len(it["r"]) and num(it["r"][c - 1])]


METRIC_KEYS = {
    "contacts": ["Contacts"], "cpd": ["Contacts per Day", "Contacts Per Day"],
    "conv": ["Conversion", "Conversion*"], "leads": ["Leads"], "lpd": ["Leads per Day"],
    "closing": ["Gross Closing", "Gross Closing*"], "gross": ["Gross Sales"], "gpd": ["Gross Sales per Day"],
    "geff": ["Gross Effectiveness"], "cxd": ["NB CXD %", "NB Net CXD Ratio"], "net": ["Net Sales"],
    "npd": ["Net Sales per Day"], "neff": ["Net Effectiveness"],
}


def ts_blocks(ts: TSSheet):
    """Group consecutive metric rows (Contacts, Leads, ...) into blocks."""
    blocks, cur = [], {}
    rev = {lab: k for k, labs in METRIC_KEYS.items() for lab in labs}
    last_row = None
    for it in ts.items:
        k = rev.get(it["label"])
        if k is None:
            continue
        if k in cur or (last_row and it["row"] - last_row > 3):
            if cur:
                blocks.append(cur)
            cur = {}
        cur[k] = it
        last_row = it["row"]
    if cur:
        blocks.append(cur)
    return [b for b in blocks if len(b) >= 3]


def ts_checks(ts: TSSheet):
    out = []
    S = ts.name
    cols = [c for c, _ in ts.month_cols]
    per = dict(ts.month_cols)

    def val(it, c):
        r = it["r"]
        return r[c - 1] if c - 1 < len(r) and num(r[c - 1]) else None

    def record(block, key, title, test, note, sev="High", tol=0.0015, rtol=None):
        it = block.get(key)
        if not it:
            return
        bad = []
        for c in cols:
            rep = val(it, c)
            exp = test(c)
            if rep is None or exp is None:
                continue
            rt = rtol(c) if rtol else 0
            if abs(rep - exp) > tol * max(1, abs(exp)) and abs(rep - exp) > 0.0006 + rt:
                bad.append((c, rep, exp))
        if bad:
            worst = max(bad, key=lambda x: abs(x[1] - x[2]) / max(abs(x[2]), 1e-9))
            months = ", ".join(str(per[c]) for c, _, _ in bad[:8]) + (" ..." if len(bad) > 8 else "")
            out.append(F(sev, "Arithmetic (time series)", S, f"{L(worst[0])}{it['row']}", f"{it['section'] or ''} > {it['label']}",
                         it["brand"], f"{title} - fails in {len(bad)} of {len(cols)} months ({months})",
                         worst[1], worst[2], note))

    def ratio(a, b, k=1.0, sub=False):
        def f(c):
            x, y = (val(a, c) if a else None), (val(b, c) if b else None)
            if x is None or y in (None, 0):
                return None
            return (1 - x / y) if sub else x / y * k
        return f

    for b in ts_blocks(ts):
        for vk, pk in (("contacts", "cpd"), ("leads", "lpd"), ("gross", "gpd"), ("net", "npd")):
            if vk in b and pk in b:
                record(b, pk, f"{b[pk]['label']} = {b[vk]['label']} / working days",
                       lambda c, v=b[vk]: (val(v, c) / ts.days[c]) if (val(v, c) is not None and num(ts.days.get(c)) and ts.days[c]) else None,
                       "Per-day figure does not equal the monthly volume divided by the working days in the header.",
                       sev="Medium", tol=0.003)
        if "gross" in b and "leads" in b:
            record(b, "closing", "Gross Closing = Gross Sales / Leads", ratio(b["gross"], b["leads"]),
                   "Closing % does not equal Gross Sales / Leads for these months.", tol=0.002,
                   rtol=lambda c, d=b["leads"]: 1.5 / val(d, c) if val(d, c) else 0)
        if "net" in b and "gross" in b:
            record(b, "cxd", "NB CXD % = 1 - Net Sales / Gross Sales", ratio(b["net"], b["gross"], sub=True),
                   "CXD % does not equal 1 - Net/Gross for these months.", tol=0.002,
                   rtol=lambda c, d=b["gross"]: 2.5 / val(d, c) if val(d, c) else 0)
        for k, lab in (("conv", "Conversion"), ("geff", "Gross Effectiveness"), ("neff", "Net Effectiveness")):
            it = b.get(k)
            if not it:
                continue
            over = [(c, val(it, c)) for c in cols if val(it, c) is not None and val(it, c) > 1.0001]
            if over:
                w = max(over, key=lambda x: x[1])
                out.append(F("High", "Sanity (time series)", S, f"{L(w[0])}{it['row']}", f"{it['section'] or ''} > {it['label']}",
                             it["brand"], f"{lab} above 100% in {len(over)} month(s): " +
                             ", ".join(str(per[c]) for c, _ in over[:8]), w[1], 1.0,
                             "Rate > 100% means numerator and denominator are from different populations."))
        for k in ("contacts", "leads", "gross", "net"):
            it = b.get(k)
            if not it:
                continue
            neg = [(c, val(it, c)) for c in cols if val(it, c) is not None and val(it, c) < 0]
            if neg:
                out.append(F("Medium", "Sanity (time series)", S, f"{L(neg[0][0])}{it['row']}", f"{it['section'] or ''} > {it['label']}",
                             it["brand"], f"Negative {it['label']} in {len(neg)} month(s)", neg[0][1], 0,
                             "Negative volume - typically net cancellations booked to a closed line."))
    # target variance columns (Act vs Tar) on every row
    sp = ts.special
    pairs = [(k, k.replace("Act", "Tar")) for k in sp if k.endswith("Act") and k.replace("Act", "Tar") in sp]
    for act_k, tar_k in pairs:
        ca, ct = sp[act_k], sp[tar_k]
        cv = next((c for c in range(ct + 1, ct + 6) if isinstance(ts.g(ts.cfg["cap_row"], c), str)
                   and "Var" in ts.g(ts.cfg["cap_row"], c)), None)
        if not cv:
            continue
        for it in ts.items:
            r = it["r"]
            a = r[ca - 1] if ca - 1 < len(r) else None
            t = r[ct - 1] if ct - 1 < len(r) else None
            v = r[cv - 1] if cv - 1 < len(r) else None
            if num(a) and num(t) and t != 0 and num(v) and abs(a / t - 1 - v) > 0.0011:
                out.append(F("High", "Arithmetic (time series)", S, f"{L(cv)}{it['row']}",
                             f"{it['section'] or ''} > {it['label']}", it["brand"],
                             f"{ts.g(ts.cfg['cap_row'], cv)} = {act_k} / {tar_k} - 1", v, a / t - 1,
                             "Variance column does not equal Actual / Target - 1."))
    return out


# =====================================================================================
#  CROSS-SHEET CHECKS
# =====================================================================================
def cross_checks(wbrows, execs, tss):
    out = []
    mom = tss.get("Month on Month")
    jul, jun, ytd = execs["Exec - F27 July"], execs["Exec - F26 June"], execs["Exec - F26 June YTD"]
    if mom:
        g = mom.g
        cap = {v: k for k, v in mom.special.items()}
        # YTD columns that include July
        for k, c in mom.special.items():
            if "YTD June" in k or "YTD Jun" in k:
                f, t = g(mom.cfg["from_row"], c), g(mom.cfg["to_row"], c)
                if num(f) and num(t) and int(t) > 202606:
                    out.append(F("High", "Period / label", "Month on Month", f"{L(c)}3:{L(c)}4", k, None,
                                 f"Column labelled '{k}' covers {int(f)} to {int(t)}", f"{int(f)}-{int(t)}",
                                 "202507-202606", "The 'F26 YTD June' column actually sums 13 months (Jul-25 to Jul-26, "
                                 "incl. projected July). YTD actual, target and variance are overstated "
                                 "(e.g. Contacts 2,985,725 vs true F26 YTD 2,736,250)."))
        # July 'Act' = projection
        if "202607 Act" in mom.special:
            c = mom.special["202607 Act"]
            v = g(11, c)
            if num(v) and num(jul.rows[6]["v"]["G"]) and v == jul.rows[6]["v"]["G"] and v != jul.rows[6]["v"]["F"]:
                out.append(F("Medium", "Period / label", "Month on Month", f"{L(c)}11", "Total > Contacts", None,
                             "'202607 Act' column holds the July PROJECTION, not the actual",
                             v, jul.rows[6]["v"]["F"],
                             "July 'Actual' in the trend sheets is the run-rate projection (Exec G6), while the "
                             "Exec sheet actual is F6. Label as 'Projected' or use actuals - otherwise July looks "
                             "~4.5% better than what has been booked."))
        # June monthly vs Exec June
        cjun = next((c for c, p in mom.month_cols if p == 202606), None)
        for r, lab, col in [(11, "Contacts", "F"), (14, "Leads", "S"), (16, "Gross Sales", "AF"), (20, "Net Sales", "AT")]:
            v = g(r, cjun) if cjun else None
            e = jun.rows[6]["v"][col]
            if num(v) and num(e) and abs(v - e) > 1:
                out.append(F("High", "Cross-sheet", "Month on Month", f"{L(cjun)}{r}", f"Total > {lab}", None,
                             "June in Month on Month = Exec - F26 June total", v, e,
                             "June figures disagree between the trend sheet and the June Exec sheet."))
    # working days for July
    days = {}
    for n, ts in tss.items():
        c = next((c for c, p in ts.month_cols if p == 202607), None)
        if c:
            days[n] = ts.days.get(c)
    summ = wbrows.get("Summary")
    if summ:
        for c, v in enumerate(summ[3], 1):
            if v == 202607 or v == "202607":
                days["Summary"] = summ[4][c - 1]
    if len({d for d in days.values() if num(d)}) > 1:
        out.append(F("Medium", "Period / label", "multiple", "working-day header rows", "July 2026 working days", None,
                     "Same month must use the same working-day count",
                     "; ".join(f"{k}: {v}" for k, v in days.items()), None,
                     "July is 23 working days in the trend sheets but 22 in Summary. Per-day rates for July "
                     "differ by ~4.5% depending on the sheet."))
    nm = tss.get("NEW - Month on Month")
    if nm:
        for k, c in nm.special.items():
            if "YTD Mar" in k:
                out.append(F("Medium", "Period / label", "NEW - Month on Month", f"{L(c)}4", k, None,
                             "Stale period column", f"{nm.g(2, c)}-{nm.g(3, c)}", "202507-202606",
                             "Fiscal-year column still shows F26 YTD March (Jul-25 to Mar-26) while the rest of the "
                             "sheet runs to Jul-26. Refresh or relabel."))
            if "YTD June" in k:
                f, t = nm.g(2, c), nm.g(3, c)
                if num(t) and int(t) > 202606:
                    out.append(F("High", "Period / label", "NEW - Month on Month", f"{L(c)}2:{L(c)}3", k, None,
                                 f"Column labelled '{k}' covers {int(f)} to {int(t)}", f"{int(f)}-{int(t)} ({nm.days.get(c)} days)",
                                 "202507-202606 (250 days)", "Same 13-month 'YTD June' issue as Month on Month."))
    return out


# =====================================================================================
#  BRAND x ITEM SHIFT ANALYSIS
# =====================================================================================
KEY_METRICS = [("Contacts", "F", "G", "H", "vol"), ("Conversion %", "L", None, "M", "rate"),
               ("Leads", "S", "T", "U", "vol"), ("Gross Closing %", "Y", None, "Z", "rate"),
               ("Gross Sales", "AF", "AG", "AH", "vol"), ("Gross Effectiveness %", "AL", None, "AM", "rate"),
               ("NB Net CXD %", "AP", None, "AQ", "rate"), ("NB Net Sales", "AT", "AU", "AV", "vol")]
DENOM = {"Conversion %": "F", "Gross Closing %": "S", "Gross Effectiveness %": "F", "NB Net CXD %": "AF"}


def brand_item_analysis(execs):
    jul, jun, ytd = execs["Exec - F27 July"], execs["Exec - F26 June"], execs["Exec - F26 June YTD"]

    def index(sh):
        d = {}
        for r in sh.rows.values():
            if r["block_brand"] and r["label"]:
                lab = "Brand Total" if r["label"].endswith("Brand Total") else r["label"]
                d.setdefault((r["block_brand"], lab), r)
            elif r["label"] and not r["block_brand"]:
                d.setdefault(("All brands", r["label"]), r)
        return d
    J, U, Y = index(jul), index(jun), index(ytd)
    rows = []
    for key, rj in J.items():
        brand, item = key
        ru, ry = U.get(key), Y.get(key)
        for m, a, p, t, kind in KEY_METRICS:
            ja, jp, jt = rj["v"][a], (rj["v"][p] if p else None), rj["v"][t]
            ua = ru["v"][a] if ru else None
            ut = ru["v"][t] if ru else None
            ya = ry["v"][a] if ry else None
            yt = ry["v"][t] if ry else None
            if not any(num(x) for x in (ja, ua, ya)):
                continue
            base = jp if kind == "vol" else ja
            jvar = (base / jt - 1) if (num(base) and num(jt) and jt) else None
            uvar = (ua / ut - 1) if (num(ua) and num(ut) and ut) else None
            yvar = (ya / yt - 1) if (num(ya) and num(yt) and yt) else None
            if kind == "vol":
                mom = (jp / ua - 1) if (num(jp) and num(ua) and ua) else None
                ytd_avg = (ya / 12) if num(ya) else None
                vs_avg = (jp / ytd_avg - 1) if (num(jp) and ytd_avg) else None
                size = max(x for x in (ja, ua, 0) if num(x))
            else:
                mom = (ja - ua) if (num(ja) and num(ua)) else None
                vs_avg = (ja - ya) if (num(ja) and num(ya)) else None
                dcol = DENOM.get(m, "F")
                size = rj["v"][dcol] if num(rj["v"][dcol]) else 0
            flags = []
            material = size >= 500
            score = 0.0
            if kind == "vol":
                if mom is not None and abs(mom) >= 0.20 and material:
                    flags.append(f"Jul (proj) vs Jun {mom:+.0%}"); score += abs(mom) / 0.20
                if vs_avg is not None and abs(vs_avg) >= 0.25 and material:
                    flags.append(f"Jul vs F26 monthly avg {vs_avg:+.0%}"); score += abs(vs_avg) / 0.25
            else:
                if mom is not None and abs(mom) >= 0.05 and material:
                    flags.append(f"Jul vs Jun {mom * 100:+.1f}pp"); score += abs(mom) / 0.05
                if vs_avg is not None and abs(vs_avg) >= 0.05 and material:
                    flags.append(f"Jul vs F26 YTD {vs_avg * 100:+.1f}pp"); score += abs(vs_avg) / 0.05
                if num(ja) and ja > 1 and m != "NB Net CXD %":
                    flags.append("rate > 100% (data issue)")
            if jvar is not None and yvar is not None and material and abs(jvar - yvar) >= 0.15:
                flags.append(f"Target var swing {yvar:+.0%} (YTD) -> {jvar:+.0%} (Jul)"); score += abs(jvar - yvar) / 0.15
            if jvar is not None and uvar is not None and material and (jvar > 0) != (uvar > 0) and abs(jvar - uvar) >= 0.10:
                flags.append("crossed the target line vs June"); score += 1
            rows.append({
                "Brand": brand, "Item / Channel": item, "Metric": m, "Type": kind,
                "Jul Actual": ja, "Jul Projecting": jp, "Jul Target": jt, "Jul Var vs Tar": jvar,
                "Jun Actual": ua, "Jun Target": ut, "Jun Var vs Tar": uvar,
                "F26 YTD Actual": ya, "F26 YTD Target": yt, "F26 YTD Var vs Tar": yvar,
                "Jul vs Jun (% vol / pp rate)": mom, "Jul vs F26 avg (% vol / pp rate)": vs_avg,
                "Base size (Jul)": size, "Clear shift?": "YES" if flags else "",
                "Shift score": round(score, 2) if flags else None,
                "Shift detail": "; ".join(flags) or None,
                "Exec July row": rj["row"],
            })
    order = {b: i for i, b in enumerate(["All brands"] + BRANDS + ["TIH Multibrand"])}
    rows.sort(key=lambda r: (order.get(r["Brand"], 99), r["Exec July row"], [k[0] for k in KEY_METRICS].index(r["Metric"])))
    return rows


def trend_shifts(tss):
    """Last month vs prior month and vs trailing 12 months, every labelled row with a brand."""
    out = []
    for name, ts in tss.items():
        for it in ts.items:
            ser = ts.series(it)
            if len(ser) < 8:
                continue
            vals = [v for _, _, v in ser]
            c_last, p_last, last = ser[-1]
            prior = vals[-2]
            trail = vals[-13:-1] if len(vals) >= 13 else vals[:-1]
            mu = st.fmean(trail)
            sd = st.pstdev(trail) if len(trail) > 2 else 0
            z = (last - mu) / sd if sd else None
            lab = it["label"].lower()
            is_rate = any(w in lab for w in ("conversion", "closing", "effectiveness", "cxd", "ratio", "distribution", "%"))
            if is_rate:
                chg, chg_avg = last - prior, last - mu
                material = abs(chg_avg) >= 0.03 or abs(chg) >= 0.03
                big = max(abs(x) for x in trail + [last]) <= 1.5
            else:
                chg = (last / prior - 1) if prior else None
                chg_avg = (last / mu - 1) if mu else None
                material = (chg is not None and abs(chg) >= 0.15) or (chg_avg is not None and abs(chg_avg) >= 0.20)
                big = abs(mu) >= 20
            # step change: last-3 mean vs prior-12 mean
            step = None
            if len(vals) >= 15:
                a, b = st.fmean(vals[-3:]), st.fmean(vals[-15:-3])
                sdb = st.pstdev(vals[-15:-3])
                step = (a - b) / sdb if sdb else None
            flag = big and material and ((z is not None and abs(z) >= 2.0) or (step is not None and abs(step) >= 2.0))
            out.append({
                "Sheet": name, "Row": it["row"], "Section": it["section"], "Brand": it["brand"] or "All / n.a.",
                "Item / Metric": it["label"], "Type": "rate" if is_rate else "volume",
                "Last month": p_last, "Last value": last, "Prior value": prior, "Trailing-12 mean": mu,
                "Trailing-12 std": sd, "Change vs prior (% vol / pp rate)": chg,
                "Change vs T12 mean (% vol / pp rate)": chg_avg, "Z-score vs T12": z,
                "Step z (last 3 vs prior 12)": step, "Clear shift?": "YES" if flag else "",
                "Note": ("July value in this sheet is the run-rate PROJECTION" if p_last == 202607 and name in
                         ("Month on Month",) else None),
            })
    return out


# =====================================================================================
def error_value_checks(wbrows):
    """Excel error codes (#DIV/0!, #VALUE!, #REF!) that were pasted into the file as values."""
    errs = {"#DIV/0!": "Medium", "#VALUE!": "High", "#REF!": "High", "#N/A": "Medium", "#NAME?": "High"}
    out = []
    for name, rows in wbrows.items():
        found = defaultdict(list)
        for i, r in enumerate(rows, 1):
            for j, v in enumerate(r, 1):
                if isinstance(v, str) and v.strip() in errs:
                    found[v.strip()].append(f"{L(j)}{i}")
        for e, cells in found.items():
            what = {"#DIV/0!": "a ratio was divided by a zero/blank volume (e.g. variance vs a 0 target, average SPV with 0 sales)",
                    "#VALUE!": "a formula hit text in a numeric cell before it was pasted as values",
                    "#REF!": "the source formula pointed at a deleted row/column/sheet before it was pasted as values - "
                             "the number that should be here is lost"}.get(e, "")
            out.append(F(errs[e], "Error values", name, ", ".join(cells[:12]) + (" ..." if len(cells) > 12 else ""),
                         f"{len(cells)} cells", None, f"{e} stored as a value", len(cells), 0,
                         f"{len(cells)} cells contain {e}: {what}. They break any SUM / AVERAGE that touches them "
                         "and show as errors in the report. Wrap source formulas in IFERROR or blank them."))
    return out


def dedupe(findings):
    """Same defect repeated in several copies of a block -> one finding, all cells listed."""
    import re
    merged = {}
    for f in findings:
        chk = re.sub(r"row \d+(-\d+)?", "row #", f["Check"])
        rep = round(f["Reported"], 6) if num(f["Reported"]) else str(f["Reported"])
        exp = round(f["Expected / Re-derived"], 6) if num(f["Expected / Re-derived"]) else str(f["Expected / Re-derived"])
        key = (f["Sheet"], f["Category"], f["Item"], f["Brand"], chk, rep, exp)
        if key in merged:
            merged[key]["Cell / Row"] += ", " + f["Cell / Row"]
            merged[key]["Occurrences"] += 1
        else:
            f = dict(f); f["Occurrences"] = 1
            merged[key] = f
    return list(merged.values())


def load_all(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = {ws.title: [tuple(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
    wb.close()
    return rows


def run(path):
    wbrows = load_all(path)
    execs = {n: ExecSheet(wbrows[n], n) for n in EXEC_SHEETS if n in wbrows}
    tss = {n: TSSheet(wbrows[n], n, cfg) for n, cfg in TS_SHEETS.items() if n in wbrows}
    findings = []
    for sh in execs.values():
        findings += exec_checks(sh)
        findings += exec_hierarchy_checks(sh)
        findings += exec_split_checks(sh)
        findings += exec_duplicate_checks(sh)
        findings += exec_brand_recon(sh)
    for ts in tss.values():
        findings += ts_checks(ts)
    findings += cross_checks(wbrows, execs, tss)
    findings += error_value_checks(wbrows)
    findings = dedupe(findings)
    sev = {"High": 0, "Medium": 1, "Low": 2, "Info": 3}
    findings.sort(key=lambda f: (sev[f["Severity"]], f["Category"], f["Sheet"]))
    for i, f in enumerate(findings, 1):
        f["ID"] = f"F{i:04d}"
    brand_rows = brand_item_analysis(execs)
    trends = trend_shifts(tss)
    return dict(findings=findings, brand=brand_rows, trends=trends, execs=execs, tss=tss, wbrows=wbrows)


if __name__ == "__main__":
    res = run(sys.argv[1])
    from collections import Counter
    print("Findings:", len(res["findings"]))
    print(Counter((f["Severity"], f["Category"]) for f in res["findings"]).most_common())
    print("Brand rows:", len(res["brand"]), "shifts:", sum(1 for r in res["brand"] if r["Clear shift?"]))
    print("Trend rows:", len(res["trends"]), "shifts:", sum(1 for r in res["trends"] if r["Clear shift?"]))
