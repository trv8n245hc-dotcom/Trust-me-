# excel_analysis

Scripts that profile and audit the **UPdate** channel-performance workbook. They work on any refreshed copy that keeps the same layout.

## Excel macro (no Python needed): `UPdateAudit.bas`

1. Open the workbook in Excel and press **Alt+F11**. Choose **File > Import File…** and pick `UPdateAudit.bas`.
2. Paste or refresh your Exec and trend sheets as usual.
3. Press **Alt+F8**, choose **RunFullAudit** and click **Run**. Save as **.xlsm** to keep the macro.

Each run deletes and rebuilds these sheets only: **Audit Summary**, **Weighted Conv & Closing**, **Calc Errors**, **Brand x Item**, **Trend Shifts**, **Meta - Sheets / Columns / Rows**. Your data sheets are never changed.

The macro finds sheets by layout, not by name, so next month's renamed sheets still work. An Exec sheet is one where C2 starts with "From Date": the latest single-month sheet is the current month, the one before it is the prior month, and the sheet where B2 ≠ B3 is YTD. A trend sheet is one with a row of YYYYMM periods near the top. The Weighted sheet uses live formulas that point at the current Exec sheet. If you paste new numbers into that sheet, the weighted figures update straight away. Re-run the macro after a layout change so the row references get refreshed.
`RunFullAuditSilent` does the same thing without pop-ups, for scheduled runs.

| Script | What it does |
|---|---|
| `profile_workbook.py` | Generic metadata for **any** `.xlsx`: one row per sheet, per used column and per non-empty row. It covers the used range, formulas, merged cells, header and period rows, monthly-column range, working days, data-type mix, error-value cells, numeric stats, the brand and section each row belongs to, and flags. |
| `analyse_update.py` | Domain checks for this workbook. It re-derives every calculated column in the three Exec sheets and checks: variances, projection factor, Closing = Sales / Leads, CXD = 1 - Net / Gross, parent = sum of children, brand pages vs the group line, repeated lines that should match, and weighted targets. It runs the same identities on the trend sheets, checks the sheets against each other (periods, working days, July projection vs actual), and finds Excel error codes pasted as values. It also builds the brand × item comparison (July vs June vs F26 YTD) and a trend-shift scan. |
| `build_report.py` | Writes a copy of the workbook with the new sheets added: READ ME, Key Findings, **Weighted Conv & Closing** (live formulas), Calc Errors, Brand x Item, Trend Shifts, Meta - Sheets / Columns / Rows. |

```bash
python excel_analysis/build_report.py  UPdate.xlsx  UPdate_Analysis.xlsx   # full deliverable
python excel_analysis/profile_workbook.py  UPdate.xlsx                     # metadata only
python excel_analysis/analyse_update.py    UPdate.xlsx                     # console summary
```

You need `openpyxl` (pandas is not required). The report formulas evaluate when the file is opened in Excel. To store cached values (for previews), run the file once through LibreOffice.

## How the weighted targets work (summary)

* **Conversion %** = eligible leads / eligible contacts. Eligible lines are the ones where a contact comes before the lead, such as OQ-PNV and Hippo-Branded.
* **Gross Closing %** = gross sales / leads.
* **Target** = target numerator / target denominator. This is the plan rate at the plan mix.
* **W Target** = Σ(actual volume_i × target rate_i) / Σ actual volume_i. This is the plan rate at the actual mix. It is built bottom-up, and a line with no target counts as 0%.
* **Mix effect** = W Target − Target. **Rate effect** = Actual − W Target. **W Tar Var** = Actual / W Target − 1.
