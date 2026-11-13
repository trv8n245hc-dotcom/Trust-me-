# excel_analysis

Scripts that profile and audit the **UPdate** channel-performance workbook. They work on any refreshed copy that keeps the same layout.

## Excel macro: `UPdateReport.bas` (use this one)

1. Open the workbook and press **Alt+F11**. Remove any old `UPdateAudit` module, then choose **File > Import File…** and pick `UPdateReport.bas`.
2. Paste your refreshed sheets.
3. Press **Alt+F8**, choose **RunReport** and click **Run**. Save as **.xlsm**. (`RunAudit` is optional and builds the replica sheets.)

Each run builds these sheets. Your data sheets are never changed.

| Sheet | What it shows |
|---|---|
| Clean - <period> (month, prior month, YTD) | Built from the Exec sheets. Includes Traditional (Generic + E-Commerce), Competition, Online Quoting (PNV + PV, without End to End), Direct Marketing and Hippo (Branded + Transfers), for Budget, FFW, Dialdirect, A&G and their 4-brand total. Virseker, Affiliates, Outbound / Multibrand and Non-Direct are excluded. Contacts, leads and sales show Actual, Projecting and Target (target summed over the included lines only), plus the variance. Conversion and effectiveness are weighted by contacts, and closing = sales ÷ leads, each for actual and target. Below the tables: a reconciliation (reported Direct = clean total + each excluded item) and a method check (the reported rates rebuilt from their lines). |
| Source x Brand | Summary-tab sources (Hippo Branded, Google, Social Media, Internet Organic / Direct, Display Banner, DM Internal, Online Quoting PV / PNV, JHB Inbound) for each brand and the 4-brand total, current vs prior month. |
| Day by Day | Every day, per brand and 4-brand total, Mondays shaded. |
| Monday View | Mon vs Tue–Thu vs Fri vs Sat, per source and brand, with sales lost per Monday. |
| Source Explorer | Pick a source and a brand from the dropdowns to see every month, week and day side by side. |
| Weighted Conv & Closing | How targets and W Targets are built. |
| Source Data | The tidy base table (links to the Summary tab) that the source views use. |

The rules are set as constants at the top of the module (`SOURCES_INCLUDED`, etc.). `RunAudit` adds the optional replica and Report Checks sheets.

**Why it stays correct on any version of the document.** No column letters or row numbers are fixed. Exec columns are found from the header text, rows from their label within each brand page, and Summary months, weeks and days from the "Period ->" and weekday rows. `tests/verify_clean.py` rebuilds every clean-view and source-view figure from the rules independently. It passed (6,591 values, 0 failures) on the July file and on a copy with columns and rows inserted, the sheet renamed and contacts changed by +37%.

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
