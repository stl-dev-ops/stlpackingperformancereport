# STLPackingPerformanceReport

Workbook location: `C:\dev\STLPackingPerformanceReport\STLPackingPerformanceReport.xlsm`

Data source:
- CERM SQL Server via trusted connection to `STL-SQL1\CRMDB`, database `sqlb00`
- `dbo.vw_stlPackingEstimatedVsActual`

Repository layout:
- `sql/packing_estimated_vs_actual.sql` is the checked-in report query used by the workbook refresh.
- `vba/` contains the imported VBA modules used to build the macro-enabled workbook.
- `scripts/build-workbook.vbs` creates a fresh `.xlsm` from the checked-in VBA source.
- `scripts/run-e2e.vbs` exercises the workbook UI automation surface and writes evidence to `artifacts/e2e`.

Configuration behavior:
- Filters live on the `Configuration` sheet and default to `NULL` semantics when left blank.
- Text filters accept SQL `LIKE` patterns such as `%ACME%`.
- Connection settings default to `STL-SQL1\CRMDB` and `sqlb00` unless overridden on the `Configuration` sheet.

Dashboard behavior:
- Shape-based controls sit across the top of the dashboard for date filters, text filters, refresh, clear, and workbook navigation.
- KPI cards summarize jobs, customers, allocated estimated hours, actual hours, variance, over-target jobs, and Variance %.
- Dashboard summary grids aggregate the live result set by work center and employee.
- KPI cards and summary sections include tooltip comments to explain each metric.
- Interaction and test instrumentation writes to dedicated `Interaction Log` and `Test Results` sheets.

Calculation rules:
- The view repeats the full job estimate on each employee/date/work-center row. The report allocates it using `EstimatedPackingMinutes * ActualShareOfJob` so it is not counted repeatedly. Estimate-only jobs retain their full estimate.
- Variance is actual minus allocated estimate. Positive means over estimate; negative means under estimate.
- Variance % is (total actual - total allocated estimate) divided by total allocated estimate. 0% is on estimate; +20% is 20% over; -20% is 20% under. No positive estimate displays `N/A`; zero actual with a positive estimate is -100%.
- The overall percentage uses ALL filtered rows, not the dashboard's 12-row preview and not an unweighted average of row, employee, or work-center percentages. Zero actual rows and actual-only hours are included.
- Each scoreboard uses the same calculation within its group. Employee estimates are proportional allocations, not independently planned employee budgets.
- Scoreboards show the dimension, distinct jobs, estimated hours, actual hours, variance hours, and Variance %, without row counts. Each data cell has a donor-style tooltip with its formula, displayed value, and contributing job IDs with per-job hours and variance percentages. Tooltips are rebuilt from all filtered detail rows on each refresh.
- Filters still limit the report. Allocation shares use lifetime job actual time, so clearing filters includes the complete job while a date/employee subset includes only its allocated portion.

Refresh process:
1. Open the workbook with macros enabled.
2. Use the top-row shape controls on the `Dashboard` sheet to adjust date and text filters, or edit values directly on `Configuration`.
3. Click `Refresh` on the `Dashboard` sheet.
4. The refresh loads `sql/packing_estimated_vs_actual.sql`, applies workbook filter values, and writes the result set to `Packing Detail`.
5. Query execution metadata is appended to the `Query Log` sheet.
6. UI actions are appended to `Interaction Log`, and automated test runs append to `Test Results`.

Build process:
- Run `cscript //nologo .\scripts\build-workbook.vbs` from `C:\dev\STLPackingPerformanceReport`.
- The script imports the checked-in VBA modules into a new workbook, runs the initializer, and saves `STLPackingPerformanceReport.xlsm`.

Automation process:
- Run `cscript //nologo .\scripts\run-e2e.vbs` from `C:\dev\STLPackingPerformanceReport`.
- The script rebuilds the workbook, opens Excel, exercises the dashboard workflows, captures PNGs of key sheets, and writes a visual layout audit CSV plus a run log to `artifacts\e2e`.

Named ranges:
- `cfgEstimateDateFrom`
- `cfgEstimateDateTo`
- `cfgDeliveryDateFrom`
- `cfgDeliveryDateTo`
- `cfgActualWorkDateFrom`
- `cfgActualWorkDateTo`
- `cfgCustomerLike`
- `cfgWorkCenterLike`
- `cfgEmployeeLike`
- `cfgServerName`
- `cfgDatabaseName`
- `cfgConnectionTimeout`
- `cfgCommandTimeout`
- `cfgWorkbookVersion`
- `cfgStatus`

Known limitations:
- The workbook is intentionally scoped to the single packing estimated-vs-actual query rather than the multi-feed production dashboard used by the donor repository.
- Filter substitution is performed by VBA against the checked-in SQL declarations, so the declaration lines in `sql/packing_estimated_vs_actual.sql` should remain present.
- Shape-based filters use InputBox prompts rather than custom calendar popup canvases, but the dashboard interaction and automation surfaces follow the donor repo pattern.
