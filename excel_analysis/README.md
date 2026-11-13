# excel_analysis

Scripts that profile and audit the **UPdate** channel-performance workbook. They work on any refreshed copy that keeps the same layout.

## Excel macro: `UPdateReport.bas` (use this one)

1. Open the workbook and press **Alt+F11**. Remove any old `UPdateAudit` module, then choose **File > Import File…** and pick `UPdateReport.bas`.
2. Paste your refreshed sheets.
3. Press **Alt+F8**, choose **RunReport** and click **Run**. Save as **.xlsm**.

The macro builds these sheets on every run. It never changes your data sheets.

| Sheet | What it is |
|---|---|
| Report Checks | Control sheet. Shows which sheets and columns were found, how many recalculated values differ, Brand Total vs Direct + Non-Direct, and every parent line whose child lines don't add up (with links). |
| Replica - <period> | One per Exec sheet: same rows, labels and metric groups, with merged headings. Volumes link to your cells. Tar Dif, Tar Var, Closing, CXD and W Tar Var are live formulas, and any cell that differs from your sheet is shaded orange. Parent-vs-children sums sit on the right. |
| Brand x Source | Every Summary-tab block by brand: Dialler totals, JHB Digital Pure, KZN, Hippo Branded, Affiliates, Google, Social Media, Internet, DM, Display, partners (JAG…), Online Quoting and JHB Inbound (Connex). Shows current vs prior month, recalculated rates, and a check against the reported conversion. |
| Monday View | The same blocks: Mon vs Tue–Thu vs Fri vs Sat, using totals (leads ÷ processed). Includes the Monday gap and sales lost per Monday. |
| Weighted Conv & Closing | How targets and W Targets are built, rebuilt with formulas. |

**Why it stays correct on any version of the document.** No column letters or row numbers are fixed. Exec columns are found from the header text ("Contacts", "Leads", … / "Actual", "Target", …), and rows from their label within their own block. Summary months, weeks and days are found from the "Period ->" and weekday rows. `tests/verify_report.py` recalculates every value independently. It passed (65,116 values, 0 failures) on the July file. It also passed on a copy with extra columns and rows inserted, the sheet renamed and contacts changed by +37%.

Colours: green is positive, red is negative, black is everything else. CXD variances are inverted, because more cancellations is bad.

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
