# Excel VBA Quality Tracker

A small Excel VBA tool that imports inspection data from a CSV file, checks every row, and builds a yield report per process step with a chart. I wrote it to practice VBA on the kind of quality and process data that comes from a semiconductor line (lots, process steps, tools, measured values and spec limits). The data in `samples/` is generated and not real.

## What it does

1. **Import** (`modImport`): reads a CSV into the sheet `Data` as plain text.
2. **Validate** (`modValidate`): checks every row in a fixed order and writes problems to the sheet `Issues`. Valid rows go to the sheet `Clean` with real numbers and dates.
   - missing LotID
   - measured value is not a number
   - spec limit is not a number, or SpecMin is greater than SpecMax
   - invalid date (format `yyyy-mm-dd`, checks leap years)
   - duplicate LotID and ProcessStep
3. **Report** (`modReport`): one line per process step with number of measurements, out of spec, yield in percent, mean, min and max. Yield below 95 percent is marked red and a column chart is added. `ExportReportPdf` saves the report as PDF.
4. **Tests** (`modTests`): unit tests for the pure helper functions in `modLogic`.

## Files

| File | Purpose |
| --- | --- |
| `vba/modLogic.bas` | pure functions: number and date parsing, spec check, yield |
| `vba/modImport.bas` | sheet helpers and CSV import |
| `vba/modValidate.bas` | row validation, Issues and Clean sheets |
| `vba/modReport.bas` | report, conditional formatting, chart, PDF export |
| `vba/modMain.bas` | entry macro `RunQualityTracker` |
| `vba/modTests.bas` | unit tests, entry macro `RunAllTests` |
| `samples/inspections.csv` | 273 generated rows, 5 of them with problems on purpose |
| `docs/expected_results.txt` | the numbers the report should show for the sample file |

## How to run

1. Open a new Excel workbook and save it as `.xlsm` (macro enabled).
2. Press `Alt + F11`, then in the VBA editor use `File > Import File...` and import the six `.bas` files from `vba/`.
3. Run `RunAllTests` first (press `F5` inside the module, or `Alt + F8`). The message box should say that all tests passed.
4. Run `RunQualityTracker`, choose `samples/inspections.csv`, and compare the sheets `Issues` and `Report` with `docs/expected_results.txt`.

Requirements: Excel for Windows (the code uses `Scripting.Dictionary`, which is not available in Excel for Mac).

### Automated build and verification on Windows

Run `./scripts/run.ps1 -Open` from PowerShell in this directory, if local script policy permits it. The script creates `QualityTracker.xlsm`, imports all six modules, runs the unit tests without message boxes, processes the sample, checks all expected report values and the chart, and tests a trailing record with a missing LotID. It refuses to overwrite an existing workbook.

Excel must allow programmatic VBA project access for this automation. If your Trust Center policy blocks it, use the manual import steps above. The runner does not change Excel security settings. A build that fails verification must not be treated as a verified workbook.

The importer accepts eight unquoted comma-separated columns with the exact sample header. Quoted CSV fields, embedded commas and multiline fields are unsupported and are rejected. Empty files and incorrect column counts are also rejected.

### Verification in this environment (2026-10-19)

Excel 16.0 launches successfully. Automated VBA import is blocked by the current Excel security configuration, so native VBA unit tests and the complete Excel workflow have not yet run. An independent PowerShell calculation confirmed 273 sample rows, 268 valid rows, all five expected issues, and every report metric in `docs/expected_results.txt`.

## Design notes

- The helper functions do not touch worksheets, so they can be tested without Excel data.
- Everything is imported as text first and converted in the validation step. That way a bad value becomes a line in `Issues` instead of a runtime error or a silently wrong number.
- The number parser always uses the dot as decimal separator, so the result does not depend on the Windows region settings.
- A row gets only the first problem found, so one bad row is counted once.

## Ideas for later

- read the spec limits from a separate table instead of the CSV
- trend chart of the mean per day
- import from a database instead of a file
