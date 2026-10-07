# Foreign-file fixtures

Workbooks written by other libraries, kept so the reader and `Workbook` are tested on files this
package did not produce. Each file holds the same small report: a header row, a data row, a total
row with a formula and a footnote row with a merged range, separated by blank rows, plus a hidden
second sheet.

| File | Written by | Notes |
| --- | --- | --- |
| `exceljs-report.xlsx` | ExcelJS 4.4.0 | Styles, whole-number validation, freeze pane, cached formula value. |
| `sheetjs-report.xlsx` | SheetJS `xlsx` 0.18.5 | The date is stored as an ISO `t="d"` cell. No styles or freeze pane. |
| `hucre-report.xlsx` | hucre 1.2.0 | Styles, freeze pane. The formula has no cached value. |

Regenerate them with `node scripts/generate-fixtures.mjs`. The script installs the pinned versions
in a temporary directory, so the output differs only in embedded timestamps. The files contain
invented data only and are free to redistribute under the repository license.

Files saved by Excel, LibreOffice and Google Sheets are still missing; add them here when they are
available, with a row in the table above.
