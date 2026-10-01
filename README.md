# Packing Estimated vs Actual Report

Workbook location: `C:\dev\STLPackingPerformanceReport\PackingEstimatedVsActualReport.xlsm`

Data source:
- CERM SQL Server via trusted connection to `STL-SQL1\CRMDB`, database `sqlb00`
- `dbo.vw_stlPackingEstimatedVsActual`

Repository layout:
- `sql/packing_estimated_vs_actual.sql` is the checked-in report query used by the workbook refresh.
- `vba/` contains the imported VBA modules used to build the macro-enabled workbook.
- `scripts/build-workbook.vbs` creates a fresh `.xlsm` from the checked-in VBA source.

Configuration behavior:
- Filters live on the `Configuration` sheet and default to `NULL` semantics when left blank.
- Text filters accept SQL `LIKE` patterns such as `%ACME%`.
- Connection settings default to `STL-SQL1\CRMDB` and `sqlb00` unless overridden on the `Configuration` sheet.

Refresh process:
1. Open the workbook with macros enabled.
2. Adjust any filters on the `Configuration` sheet.
3. Click `Refresh` on the `Dashboard` sheet.
4. The refresh loads `sql/packing_estimated_vs_actual.sql`, applies workbook filter values, and writes the result set to `Packing Detail`.
5. Query execution metadata is appended to the `Query Log` sheet.

Build process:
- Run `cscript //nologo .\scripts\build-workbook.vbs` from `C:\dev\STLPackingPerformanceReport`.
- The script imports the checked-in VBA modules into a new workbook, runs the initializer, and saves `PackingEstimatedVsActualReport.xlsm`.

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
