On Error Resume Next

Dim shell
Dim fso
Dim xl
Dim wb
Dim wsDash
Dim runId
Dim artifactsRoot
Dim exportRoot
Dim logPath
Dim layoutAuditPath
Dim builtWorkbookPath
Dim dialogWatchCommand
Dim failureCount
Dim tooltipCell
Dim standaloneFolder
Dim standaloneWorkbookPath
Dim sqlSourceStream
Dim sqlSourceText
Dim standaloneFilesOnly
Dim scoreboardScrollRow
Dim dynamicGridResult

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

artifactsRoot = "C:\dev\STLPackingPerformanceReport\artifacts"
exportRoot = artifactsRoot & "\e2e"
logPath = exportRoot & "\run-e2e.log"
layoutAuditPath = exportRoot & "\dashboard-layout-audit.csv"
builtWorkbookPath = "C:\dev\STLPackingPerformanceReport\STLPackingPerformanceReport.xlsm"
dialogWatchCommand = "powershell -NoProfile -ExecutionPolicy Bypass -File ""C:\dev\STLPackingPerformanceReport\scripts\watch-excel-dialogs.ps1"" -OutputPath ""C:\dev\STLPackingPerformanceReport\artifacts\e2e\dialog-watch.log"" -TimeoutSeconds 240"

EnsureFolder artifactsRoot
EnsureFolder exportRoot
WriteLog "Starting build"
shell.Run dialogWatchCommand, 0, False

If shell.Run("cscript //nologo ""C:\dev\STLPackingPerformanceReport\scripts\build-workbook.vbs""", 0, True) <> 0 Then
    WriteLog "Build failed"
    WScript.Echo "FAIL:BUILD"
    WScript.Quit 1
End If

WriteLog "Opening workbook"
standaloneFolder = fso.BuildPath(fso.GetSpecialFolder(2), "packing-report-" & fso.GetTempName)
EnsureFolder standaloneFolder
standaloneWorkbookPath = fso.BuildPath(standaloneFolder, fso.GetFileName(builtWorkbookPath))
Err.Clear
fso.CopyFile builtWorkbookPath, standaloneWorkbookPath, True
If Err.Number <> 0 Then
    WriteLog "FAIL | Standalone distribution | Copy failed | " & Err.Description
    WScript.Quit 1
End If
standaloneFilesOnly = fso.GetFolder(standaloneFolder).Files.Count = 1 And fso.GetFolder(standaloneFolder).SubFolders.Count = 0
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
Err.Clear
Set wb = xl.Workbooks.Open(standaloneWorkbookPath)
If Err.Number <> 0 Then
    WriteLog "FAIL | Standalone distribution | Open failed | " & Err.Description
    xl.Quit
    WScript.Quit 1
End If
Set wsDash = wb.Worksheets("Dashboard")
xl.WindowState = -4137

runId = InvokeAction(xl, wb, "new-run-id", "e2e", "")
WriteLog "Run ID: " & runId

RecordPass wb, runId, "Workbook opens", "Open workbook and initialize", "Workbook opened for automation"
AssertTrue wb, runId, "Standalone distribution", "Workbook is the only distributed file", standaloneFilesOnly And Not fso.FileExists(standaloneFolder & "\packing_estimated_vs_actual.sql") And Not fso.FileExists(standaloneFolder & "\sql\packing_estimated_vs_actual.sql"), "Before Excel creates its lock file, the folder contains only the workbook; all refresh and picker tests run without SQL sidecars"
Set sqlSourceStream = fso.OpenTextFile("C:\dev\STLPackingPerformanceReport\packing_estimated_vs_actual.sql", 1)
sqlSourceText = sqlSourceStream.ReadAll
sqlSourceStream.Close
AssertTrue wb, runId, "Standalone distribution", "Embedded SQL matches source", NormalizeSqlText(CStr(xl.Run(WorkbookMacro(wb, "modEmbeddedSql.PackingSqlText")))) = NormalizeSqlText(sqlSourceText), "Build must embed the authoritative source SQL without drift"
AssertTrue wb, runId, "Branding", "Workbook filename matches", wb.Name = "STLPackingPerformanceReport.xlsm", "Workbook should use the new filename"
AssertTrue wb, runId, "Branding", "Dashboard title matches", CStr(wsDash.Range("B2").Value) = "Packing Performance Report", "Sheet title bar should use the display title"
AssertTrue wb, runId, "Branding", "Current variance build loaded", wsDash.ListObjects("tblEmployeeSummary").ListColumns.Count = 6 And wsDash.ListObjects("tblEmployeeSummary").HeaderRowRange.Cells(1, 6).Value = "Variance %", "Six columns with variance percentage and no row count"
AssertTrue wb, runId, "Percentage edge cases", "Zero actual is minus 100 percent", CDbl(xl.Run(WorkbookMacro(wb, "modReportMain.PackingPercentage"), 0, 60)) = -1, "Zero actual must be 100 percent under estimate"
AssertTrue wb, runId, "Percentage edge cases", "Under estimate is negative", CDbl(xl.Run(WorkbookMacro(wb, "modReportMain.PackingPercentage"), 45, 60)) = -0.25, "45 actual versus 60 estimated is -25 percent"
AssertTrue wb, runId, "Percentage edge cases", "Over estimate is positive", CDbl(xl.Run(WorkbookMacro(wb, "modReportMain.PackingPercentage"), 75, 60)) = 0.25, "75 actual versus 60 estimated is +25 percent"
AssertTrue wb, runId, "Percentage edge cases", "On estimate is zero", CDbl(xl.Run(WorkbookMacro(wb, "modReportMain.PackingPercentage"), 60, 60)) = 0, "On estimate is 0 percent"
AssertTrue wb, runId, "Percentage edge cases", "Zero estimate is N/A", CStr(xl.Run(WorkbookMacro(wb, "modReportMain.PackingPercentage"), 60, 0)) = "N/A", "Cannot divide by zero or display a misleading zero percent"
AssertContains wb, runId, "Startup refresh", "Refresh completes automatically on open", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestDashboardStatus"))), "Refresh complete:", "Workbook_Open must refresh without an explicit automation refresh action"
AssertTrue wb, runId, "Startup refresh", "Startup query is logged", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Query Log", "tblQueryLog")) > 0, "Startup must execute and log a query"
AssertTrue wb, runId, "Startup refresh", "Startup dashboard grid is populated", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Dashboard", "tblDashboardDetail")) > 0, "Default Shipping and last-week filters should populate the dashboard on open"
AssertTrue wb, runId, "Startup refresh", "Startup grid remains aligned", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestDashboardGridAligned"))), "Startup refresh must preserve merged KPI/grid boundaries"
AssertContains wb, runId, "Workbook opens", "Dashboard sheet active on open", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCurrentSheetName"))), "Dashboard", "Workbook should open on Dashboard"
AssertContains wb, runId, "Workbook opens", "Dashboard top cell selected on open", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestActiveCellAddress"))), "B2", "Workbook should open at the top of Dashboard"
AssertTrue wb, runId, "Workbook opens", "Dashboard scroll row reset on open", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestActiveWindowScrollRow"))) = 1, "Dashboard should open scrolled to row 1"
AssertTrue wb, runId, "Workbook opens", "Dashboard scroll column reset on open", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestActiveWindowScrollColumn"))) = 1, "Dashboard should open scrolled to column 1"
dynamicGridResult = False
Err.Clear
dynamicGridResult = xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestDashboardGridResize"))
If Err.Number <> 0 Then
    RecordFail wb, runId, "Dynamic grid", "Resize macro executes", Err.Description
Else
    AssertTrue wb, runId, "Dynamic grid", "Grid grows shrinks and handles empty results", CBool(dynamicGridResult), "17, 3, 0, and 2 rows must resize exactly, retain alignment, and clear old rows, titles, and tooltips"
End If
RecordActionResult wb, runId, "Initialize", "InitializeReportWorkbook", InvokeAction(xl, wb, "initialize", "", ""), "Workbook initializer completed"
AssertContains wb, runId, "Initialize", "Dashboard sheet active after initialize", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCurrentSheetName"))), "Dashboard", "Initializer should leave workbook on Dashboard"
AssertContains wb, runId, "Initialize", "Dashboard top cell selected after initialize", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestActiveCellAddress"))), "B2", "Initializer should leave Dashboard focused at the top"

AssertTrue wb, runId, "Dashboard controls", "Estimate-from button exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "ppr_btnEstimateDateFrom")), "Estimate-from button should render"
AssertTrue wb, runId, "Dashboard controls", "Refresh button exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "ppr_btnRefresh")), "Refresh button should render"
AssertTrue wb, runId, "Dashboard controls", "KPI comment exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCommentExists"), "Dashboard", "B11")), "KPI tooltip comment should exist"
AssertTrue wb, runId, "Dashboard controls", "Customer button exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "ppr_btnCustomerFilter")), "Customer picker button should render"
AssertContains wb, runId, "Dashboard defaults", "Actual-from button shows default date", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeText"), "Dashboard", "ppr_btnActualDateFrom")), "Act From", "Dashboard should render actual date defaults on the button"
AssertTrue wb, runId, "Dashboard defaults", "Actual-from button is not Any", InStr(1, CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeText"), "Dashboard", "ppr_btnActualDateFrom")), "Any", vbTextCompare) = 0, "Dashboard should default actual-from away from Any"
AssertContains wb, runId, "Dashboard defaults", "Work center default button caption", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeText"), "Dashboard", "ppr_btnWorkCenterFilter")), "3 Work Centers", "Dashboard should default to shipping work centers"

RecordActionResult wb, runId, "Date picker", "Open estimate-from picker", InvokeAction(xl, wb, "open-estimate-from", "", ""), "Opened donor-style calendar"
AssertTrue wb, runId, "Date picker", "Calendar popup created", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprCal_bg")), "Calendar should render"
RecordActionResult wb, runId, "Date picker", "Calendar next month", InvokeAction(xl, wb, "calendar-next-month", "", ""), "Moved calendar forward"
RecordActionResult wb, runId, "Date picker", "Calendar previous month", InvokeAction(xl, wb, "calendar-prev-month", "", ""), "Moved calendar backward"
RecordActionResult wb, runId, "Date picker", "Calendar close", InvokeAction(xl, wb, "calendar-close", "", ""), "Closed calendar"
AssertTrue wb, runId, "Date picker", "Calendar popup removed", Not CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprCal_bg")), "Calendar should close cleanly"
RecordActionResult wb, runId, "Date picker", "Open delivery-from picker", InvokeAction(xl, wb, "open-delivery-from", "", ""), "Opened delivery-from calendar"
AssertTrue wb, runId, "Date picker", "Delivery calendar popup created", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprCal_bg")), "Delivery calendar should render"
RecordActionResult wb, runId, "Date picker", "Close delivery-from picker", InvokeAction(xl, wb, "calendar-close", "", ""), "Closed delivery calendar"
RecordActionResult wb, runId, "Date picker", "Open actual-to picker", InvokeAction(xl, wb, "open-actual-to", "", ""), "Opened actual-to calendar"
AssertTrue wb, runId, "Date picker", "Actual-to calendar popup created", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprCal_bg")), "Actual-to calendar should render"
RecordActionResult wb, runId, "Date picker", "Close actual-to picker", InvokeAction(xl, wb, "calendar-close", "", ""), "Closed actual-to calendar"
AssertTrue wb, runId, "Date picker", "Calendar popup removed after additional flows", Not CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprCal_bg")), "Calendar popup should be removed"

RecordActionResult wb, runId, "Value picker", "Open customer picker", InvokeAction(xl, wb, "open-customer", "", ""), "Opened customer picker"
AssertTrue wb, runId, "Value picker", "Customer popup created", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprVal_bg")), "Customer picker should render"
RecordActionResult wb, runId, "Value picker", "Toggle first customer", InvokeAction(xl, wb, "value-picker-select-first", "", ""), "Selected first visible customer"
RecordActionResult wb, runId, "Value picker", "Apply customer picker", InvokeAction(xl, wb, "value-picker-apply", "", ""), "Applied customer selection"
AssertTrue wb, runId, "Value picker", "Customer filter config updated", Len(CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestConfigText"), "cfgCustomerLike"))) > 0, "Customer picker should set a filter value"

RecordActionResult wb, runId, "Value picker", "Open employee picker", InvokeAction(xl, wb, "open-employee", "", ""), "Opened employee picker"
AssertTrue wb, runId, "Value picker", "Employee popup created", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprVal_bg")), "Employee picker should render"
RecordActionResult wb, runId, "Value picker", "Close employee picker", InvokeAction(xl, wb, "value-picker-close", "", ""), "Closed employee picker"
AssertTrue wb, runId, "Value picker", "Employee popup removed", Not CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprVal_bg")), "Employee picker should close cleanly"
RecordActionResult wb, runId, "Value picker", "Open work center picker", InvokeAction(xl, wb, "open-workcenter", "", ""), "Opened work center picker"
AssertTrue wb, runId, "Value picker", "Work center popup created", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprVal_bg")), "Work center picker should render"
RecordActionResult wb, runId, "Value picker", "Close work center picker", InvokeAction(xl, wb, "value-picker-close", "", ""), "Closed work center picker"
AssertTrue wb, runId, "Value picker", "Work center popup removed", Not CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprVal_bg")), "Work center picker should close cleanly"

RecordActionResult wb, runId, "Navigation", "Go to configuration", InvokeAction(xl, wb, "goto-configuration", "", ""), "Navigate to Configuration"
AssertContains wb, runId, "Navigation", "Configuration sheet active", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCurrentSheetName"))), "Configuration", "Should land on Configuration"
RecordActionResult wb, runId, "Navigation", "Go to detail", InvokeAction(xl, wb, "goto-detail", "", ""), "Navigate to detail"
AssertContains wb, runId, "Navigation", "Detail sheet active", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCurrentSheetName"))), "Packing Detail", "Should land on Packing Detail"
RecordActionResult wb, runId, "Navigation", "Go to query log", InvokeAction(xl, wb, "goto-query-log", "", ""), "Navigate to query log"
AssertContains wb, runId, "Navigation", "Query log sheet active", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCurrentSheetName"))), "Query Log", "Should land on Query Log"
RecordActionResult wb, runId, "Navigation", "Go to dashboard", InvokeAction(xl, wb, "goto-dashboard", "", ""), "Navigate back to dashboard"

RecordActionResult wb, runId, "Filters", "Clear filters", InvokeAction(xl, wb, "clear-filters", "", ""), "Clear all filters"
RecordActionResult wb, runId, "Filters", "Set estimate range", InvokeAction(xl, wb, "set-estimate-range", "2026-01-01", "2026-12-31"), "Estimate date range applied"
RecordActionResult wb, runId, "Filters", "Set delivery range", InvokeAction(xl, wb, "set-delivery-range", "2026-01-01", "2026-12-31"), "Delivery date range applied"
RecordActionResult wb, runId, "Filters", "Set actual range", InvokeAction(xl, wb, "set-actual-range", "2026-01-01", "2026-12-31"), "Actual work date range applied"
RecordActionResult wb, runId, "Filters", "Set customer filter", InvokeAction(xl, wb, "set-customer", "%", ""), "Customer filter applied"
RecordActionResult wb, runId, "Filters", "Set work center filter", InvokeAction(xl, wb, "set-workcenter", "%", ""), "Work center filter applied"
RecordActionResult wb, runId, "Filters", "Set employee filter", InvokeAction(xl, wb, "set-employee", "%", ""), "Employee filter applied"

RecordActionResult wb, runId, "Refresh", "Refresh report", InvokeAction(xl, wb, "refresh", "", ""), CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestDashboardStatus")))
AssertTrue wb, runId, "Refresh", "Packing detail rows present", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Packing Detail", "tblPackingDetail")) > 0, "Full-range refresh should load detail rows"
AssertTrue wb, runId, "Refresh", "Dashboard detail rows present", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Dashboard", "tblDashboardDetail")) > 0, "Dashboard data grid should populate"
AssertTrue wb, runId, "Refresh", "Grid cells match KPI boundaries", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestDashboardGridAligned"))), "All seven headers and data columns must share KPI left edges and widths"
AssertPackingMath wb, runId, "Full-range math"

RecordActionResult wb, runId, "Picker workflow", "Clear filters for subset workflow", InvokeAction(xl, wb, "clear-filters", "", ""), "Reset filter state"
AssertContains wb, runId, "Picker workflow", "Work center filter reset to defaults", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestConfigText"), "cfgWorkCenterLike")), "Shipping", "Work center defaults should be restored"
AssertContains wb, runId, "Picker workflow", "Work center button reset to defaults", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeText"), "Dashboard", "ppr_btnWorkCenterFilter")), "3 Work Centers", "Work center button should reflect restored defaults"
RecordActionResult wb, runId, "Picker workflow", "Set actual week for subset workflow", InvokeAction(xl, wb, "set-actual-range", "2026-09-20", "2026-09-26"), "Applied known row-producing week"
RecordActionResult wb, runId, "Picker workflow", "Open work center picker", InvokeAction(xl, wb, "open-workcenter", "", ""), "Opened work center picker"
AssertTrue wb, runId, "Picker workflow", "Work center popup created", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "pprVal_bg")), "Work center picker should render"
RecordActionResult wb, runId, "Picker workflow", "Filter work center popup to Shipping", InvokeAction(xl, wb, "value-picker-set-filter", "Shipping", ""), "Filtered visible work centers"
RecordActionResult wb, runId, "Picker workflow", "Clear work center selections", InvokeAction(xl, wb, "value-picker-select-all", "", ""), "Cleared existing selections in picker"
RecordActionResult wb, runId, "Picker workflow", "Select visible Shipping work centers", InvokeAction(xl, wb, "value-picker-select-all-visible", "", ""), "Selected all visible work centers"
RecordActionResult wb, runId, "Picker workflow", "Apply Shipping work centers", InvokeAction(xl, wb, "value-picker-apply", "", ""), "Applied work center subset"
AssertContains wb, runId, "Picker workflow", "Work center config contains Shipping", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestConfigText"), "cfgWorkCenterLike")), "Shipping", "Picker should store selected work centers"
RecordActionResult wb, runId, "Picker workflow", "Refresh with Shipping work centers", InvokeAction(xl, wb, "refresh", "", ""), CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestDashboardStatus")))
AssertTrue wb, runId, "Picker workflow", "Shipping subset returns rows", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Packing Detail", "tblPackingDetail")) > 0, "Shipping work centers in the known week should return detail rows"
AssertTrue wb, runId, "Picker workflow", "Distinct jobs KPI updates for Shipping subset", CDbl(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestKpiNumericValue"), "B11")) > 0, "Distinct jobs KPI should be non-zero for Shipping subset"
AssertTrue wb, runId, "Picker workflow", "Actual hours KPI updates for Shipping subset", CDbl(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestKpiNumericValue"), "H11")) > 0, "Actual hours KPI should be non-zero for Shipping subset"
AssertTrue wb, runId, "Picker workflow", "Dashboard detail subset rows present", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Dashboard", "tblDashboardDetail")) > 0, "Dashboard detail grid should populate for Shipping subset"
AssertTrue wb, runId, "Picker workflow", "Subset grid cells match KPI boundaries", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestDashboardGridAligned"))), "Merged grid cells must remain aligned after subset refresh"
AssertTrue wb, runId, "Picker workflow", "Work center summary has rows", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Dashboard", "tblWorkCenterSummary")) > 0, "Work center summary should populate for Shipping subset"
AssertTrue wb, runId, "Picker workflow", "Employee summary has rows", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Dashboard", "tblEmployeeSummary")) > 0, "Employee summary should populate for Shipping subset"
AssertTrue wb, runId, "Scoreboard headers", "Work center header contrast", wsDash.ListObjects("tblWorkCenterSummary").HeaderRowRange.Font.Color = RGB(255, 255, 255) And wsDash.ListObjects("tblWorkCenterSummary").HeaderRowRange.Interior.Color = wsDash.Range("B2").Interior.Color, "White header text must have the solid dark title background"
AssertTrue wb, runId, "Scoreboard headers", "Employee header contrast", wsDash.ListObjects("tblEmployeeSummary").HeaderRowRange.Font.Color = RGB(255, 255, 255) And wsDash.ListObjects("tblEmployeeSummary").HeaderRowRange.Interior.Color = wsDash.Range("B2").Interior.Color, "White header text must have the solid dark title background"
AssertPackingMath wb, runId, "Shipping subset math"

RecordActionResult wb, runId, "Reset workflow", "Clear filters after subset workflow", InvokeAction(xl, wb, "clear-filters", "", ""), "Returned dashboard to unfiltered state"
AssertContains wb, runId, "Reset workflow", "Work center button reset", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeText"), "Dashboard", "ppr_btnWorkCenterFilter")), "3 Work Centers", "Work center button should show restored default work center filter"
AssertContains wb, runId, "Reset workflow", "Actual date defaults restored", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeText"), "Dashboard", "ppr_btnActualDateFrom")), "Act From", "Actual date default should remain visible on the button"
RecordActionResult wb, runId, "Reset workflow", "Refresh restored defaults", InvokeAction(xl, wb, "refresh", "", ""), "Saved data must match the restored default filter captions"
AssertPackingMath wb, runId, "Default-filter math"
RecordActionResult wb, runId, "Reset workflow", "Return to dashboard before save", InvokeAction(xl, wb, "goto-dashboard", "", ""), "Ensure workbook ends on dashboard"
AssertContains wb, runId, "Reset workflow", "Dashboard top cell selected before save", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestActiveCellAddress"))), "B2", "Dashboard should end focused at the top"

wb.Save
CaptureSheetWindowPng xl, wb.Worksheets("Dashboard"), exportRoot & "\dashboard.png", 1
scoreboardScrollRow = wsDash.ListObjects("tblEmployeeSummary").Range.Row - 2
CaptureSheetWindowPng xl, wb.Worksheets("Dashboard"), exportRoot & "\scoreboards.png", scoreboardScrollRow
Set tooltipCell = wsDash.ListObjects("tblEmployeeSummary").DataBodyRange.Cells(1, 6)
tooltipCell.Comment.Shape.Left = wsDash.Cells(scoreboardScrollRow + 7, 10).Left
tooltipCell.Comment.Shape.Top = wsDash.Cells(scoreboardScrollRow + 7, 10).Top
tooltipCell.Comment.Visible = True
CaptureSheetWindowPng xl, wb.Worksheets("Dashboard"), exportRoot & "\scoreboard-tooltip.png", scoreboardScrollRow
tooltipCell.Comment.Visible = False
CaptureSheetWindowPng xl, wb.Worksheets("Packing Detail"), exportRoot & "\packing-detail.png", 1
CaptureSheetWindowPng xl, wb.Worksheets("Query Log"), exportRoot & "\query-log.png", 1
WriteDashboardLayoutAudit wb.Worksheets("Dashboard"), layoutAuditPath, wb, runId
AssertTrue wb, runId, "Visual audit", "Dashboard layout issue count is zero", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestDashboardLayoutIssueCount"))) = 0, "Dashboard layout should have no audited overlaps"

AssertTrue wb, runId, "Logging", "Interaction log populated", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Interaction Log", "tblInteractionLog")) > 0, "Interaction log should have entries"
AssertTrue wb, runId, "Logging", "Query log populated", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Query Log", "tblQueryLog")) > 0, "Query log should have entries"
AssertTrue wb, runId, "Logging", "Test results populated", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Test Results", "tblTestResults")) > 0, "Test results should have entries"

RecordActionResult wb, runId, "Save workflow", "Restore dashboard after screenshots", InvokeAction(xl, wb, "goto-dashboard", "", ""), "Final saved view must be Dashboard at the top"
wb.Save
wb.Close False
xl.Quit
If failureCount = 0 Then
    Err.Clear
    fso.CopyFile standaloneWorkbookPath, builtWorkbookPath, True
    If Err.Number <> 0 Then
        failureCount = failureCount + 1
        WriteLog "FAIL | Standalone distribution | Copy verified workbook back failed | " & Err.Description
    End If
End If
fso.DeleteFolder standaloneFolder, True

WriteLog "E2E completed"
If failureCount > 0 Then
    WScript.Echo "FAIL:" & failureCount & " assertions; see " & logPath
    WScript.Quit 1
End If
WScript.Echo "OK:" & exportRoot

Function NormalizeSqlText(sqlText)
    sqlText = Replace(sqlText, vbCr, "")
    Do While Right(sqlText, 1) = vbLf
        sqlText = Left(sqlText, Len(sqlText) - 1)
    Loop
    NormalizeSqlText = sqlText
End Function

Function InvokeAction(excelApp, workbookObj, actionName, arg1, arg2)
    On Error Resume Next
    Err.Clear
    InvokeAction = CStr(excelApp.Run(WorkbookMacro(workbookObj, "modAutomationSupport.AutomationInvoke"), actionName, arg1, arg2))
    If Err.Number <> 0 Then InvokeAction = "ERR:" & Err.Number & ":" & Err.Description
End Function

Sub AssertPackingMath(workbookObj, currentRunId, scenarioName)
    Dim detailTable, summaryTable, dashboard, rowIndex, summaryIndex
    Dim estimated, actual, variance, groupEstimated, groupActual
    Dim groupKey, dimensionColumn, tableName, percentValue
    Dim detailValues, summaryValues, detailValid
    Set dashboard = workbookObj.Worksheets("Dashboard")
    Set detailTable = workbookObj.Worksheets("Packing Detail").ListObjects("tblPackingDetail")
    detailValues = detailTable.DataBodyRange.Value
    AssertTrue workbookObj, currentRunId, scenarioName, "Dashboard contains every detail row", workbookObj.Names("tblDashboardDetail").RefersToRange.Rows.Count - 1 = detailTable.ListRows.Count, "No 12-row cap or padded rows"
    AssertTrue workbookObj, currentRunId, scenarioName, "Scoreboards follow complete grid", dashboard.ListObjects("tblWorkCenterSummary").Range.Row = 19 + detailTable.ListRows.Count And dashboard.ListObjects("tblEmployeeSummary").Range.Row = 19 + detailTable.ListRows.Count, "Scoreboards must move below all result rows"
    detailValid = True
    estimated = 0
    actual = 0
    For rowIndex = 1 To UBound(detailValues, 1)
        estimated = estimated + MathNumber(detailValues(rowIndex, 10))
        actual = actual + MathNumber(detailValues(rowIndex, 11))
        variance = MathNumber(detailValues(rowIndex, 14))
        If Abs(variance - (MathNumber(detailValues(rowIndex, 11)) - MathNumber(detailValues(rowIndex, 10)))) >= 0.001 Then detailValid = False
    Next
    AssertTrue workbookObj, currentRunId, scenarioName, "Every detail variance reconciles", detailValid, "Row variance must equal row actual minus allocated estimate across all detail rows"
    AssertTrue workbookObj, currentRunId, scenarioName, "KPI estimate covers all rows", Abs(CDbl(xl.Run(WorkbookMacro(workbookObj, "modAutomationSupport.TestKpiNumericValue"), "F11")) - estimated / 60) < 0.011, "Estimate must use the full detail dataset"
    AssertTrue workbookObj, currentRunId, scenarioName, "KPI actual covers all rows", Abs(CDbl(xl.Run(WorkbookMacro(workbookObj, "modAutomationSupport.TestKpiNumericValue"), "H11")) - actual / 60) < 0.011, "Actual must use the full detail dataset"
    AssertTrue workbookObj, currentRunId, scenarioName, "KPI variance reconciles", Abs(CDbl(xl.Run(WorkbookMacro(workbookObj, "modAutomationSupport.TestKpiNumericValue"), "J11")) - (actual - estimated) / 60) < 0.011, "Variance = total actual - total allocated estimate"
    If estimated > 0 Then
        AssertTrue workbookObj, currentRunId, scenarioName, "KPI weighted percentage", Abs(CDbl(xl.Run(WorkbookMacro(workbookObj, "modAutomationSupport.TestKpiNumericValue"), "N11")) - (actual - estimated) / estimated * 100) < 0.011, "Variance percentage = all-row variance / all-row estimate, not average row percentages"
    End If
    For Each tableName In Array("tblWorkCenterSummary", "tblEmployeeSummary")
        Set summaryTable = dashboard.ListObjects(tableName)
        summaryValues = summaryTable.DataBodyRange.Value
        dimensionColumn = 8
        If tableName = "tblEmployeeSummary" Then dimensionColumn = 9
        For summaryIndex = 1 To UBound(summaryValues, 1)
            groupKey = CStr(summaryValues(summaryIndex, 1))
            groupEstimated = 0
            groupActual = 0
            For rowIndex = 1 To UBound(detailValues, 1)
                If CStr(detailValues(rowIndex, dimensionColumn)) = groupKey Or (groupKey = "(Blank)" And Len(CStr(detailValues(rowIndex, dimensionColumn))) = 0) Then
                    groupEstimated = groupEstimated + MathNumber(detailValues(rowIndex, 10)) / 60
                    groupActual = groupActual + MathNumber(detailValues(rowIndex, 11)) / 60
                End If
            Next
            AssertTrue workbookObj, currentRunId, scenarioName, tableName & " totals " & groupKey, Abs(CDbl(summaryValues(summaryIndex, 3)) - groupEstimated) < 0.001 And Abs(CDbl(summaryValues(summaryIndex, 4)) - groupActual) < 0.001 And Abs(CDbl(summaryValues(summaryIndex, 5)) - (groupActual - groupEstimated)) < 0.001, "Group hours and variance must reconcile to detail"
            percentValue = summaryValues(summaryIndex, 6)
            If groupEstimated > 0 Then
                AssertTrue workbookObj, currentRunId, scenarioName, tableName & " percentage " & groupKey, Abs(CDbl(percentValue) - (groupActual - groupEstimated) / groupEstimated) < 0.000001 And summaryTable.DataBodyRange.Cells(summaryIndex, 6).NumberFormat = "+0.00%;-0.00%;0.00%", "Group weighted variance percentage with signed formatting"
            Else
                AssertTrue workbookObj, currentRunId, scenarioName, tableName & " no estimate " & groupKey, CStr(percentValue) = "N/A", "Missing estimate must not show a misleading zero percent"
            End If
        Next
    Next
    AssertSummaryJobTooltips workbookObj, currentRunId, scenarioName, detailValues
End Sub

Sub AssertSummaryJobTooltips(workbookObj, currentRunId, scenarioName, detailValues)
    Dim tableName, summaryTable, summaryValues, dimensionColumn, summaryRow, dataRow, columnIndex
    Dim groupKey, expectedJobs, jobId, jobKey, jobHours, tooltipText, auditParts, auditLines, auditRow, auditJob
    Dim jobListValid, commentsValid, auditText, seenJobs
    For Each tableName In Array("tblWorkCenterSummary", "tblEmployeeSummary")
        Set summaryTable = workbookObj.Worksheets("Dashboard").ListObjects(tableName)
        summaryValues = summaryTable.DataBodyRange.Value
        dimensionColumn = 8
        If tableName = "tblEmployeeSummary" Then dimensionColumn = 9
        AssertTrue workbookObj, currentRunId, scenarioName, tableName & " no row count", summaryTable.ListColumns.Count = 6 And CStr(summaryTable.HeaderRowRange.Cells(1, 2).Value) = "Jobs", "Scoreboard must omit row count"
        For summaryRow = 1 To UBound(summaryValues, 1)
            groupKey = CStr(summaryValues(summaryRow, 1))
            Set expectedJobs = CreateObject("Scripting.Dictionary")
            For dataRow = 1 To UBound(detailValues, 1)
                If CStr(detailValues(dataRow, dimensionColumn)) = groupKey Or (groupKey = "(Blank)" And Len(CStr(detailValues(dataRow, dimensionColumn))) = 0) Then
                    jobId = CStr(detailValues(dataRow, 1))
                    If Len(jobId) = 0 Then jobId = "(No Job ID)"
                    jobHours = Array(0, 0)
                    If expectedJobs.Exists(jobId) Then jobHours = expectedJobs(jobId)
                    jobHours(0) = jobHours(0) + MathNumber(detailValues(dataRow, 10)) / 60
                    jobHours(1) = jobHours(1) + MathNumber(detailValues(dataRow, 11)) / 60
                    expectedJobs(jobId) = jobHours
                End If
            Next
            commentsValid = True
            auditText = ""
            For columnIndex = 1 To 6
                tooltipText = ""
                Err.Clear
                tooltipText = summaryTable.DataBodyRange.Cells(summaryRow, columnIndex).Comment.Text
                If Err.Number <> 0 Or InStr(tooltipText, "Formula:") = 0 Or InStr(tooltipText, "Displayed value:") = 0 Then commentsValid = False
                auditParts = Split(tooltipText, "Jobs included (active filters):" & vbLf)
                If UBound(auditParts) <> 1 Then
                    commentsValid = False
                Else
                    If columnIndex = 1 Then auditText = auditParts(1)
                    If auditParts(1) <> auditText Then commentsValid = False
                End If
            Next
            jobListValid = commentsValid
            Set seenJobs = CreateObject("Scripting.Dictionary")
            auditLines = Split(auditText, vbLf)
            For auditRow = 1 To UBound(auditLines)
                auditJob = Trim(Left(auditLines(auditRow), 14))
                If Not expectedJobs.Exists(auditJob) Or seenJobs.Exists(auditJob) Then
                    jobListValid = False
                Else
                    seenJobs(auditJob) = True
                    jobHours = expectedJobs(auditJob)
                    If Abs(CDbl(Trim(Mid(auditLines(auditRow), 15, 10))) - jobHours(0)) >= 0.011 Then jobListValid = False
                    If Abs(CDbl(Trim(Mid(auditLines(auditRow), 26, 10))) - jobHours(1)) >= 0.011 Then jobListValid = False
                    If Abs(CDbl(Trim(Mid(auditLines(auditRow), 37, 10))) - (jobHours(1) - jobHours(0))) >= 0.011 Then jobListValid = False
                End If
            Next
            If seenJobs.Count <> expectedJobs.Count Then jobListValid = False
            AssertTrue workbookObj, currentRunId, scenarioName, tableName & " job tooltips " & groupKey, jobListValid, "All six cells must list exactly the filtered group's jobs once, with reconcilable per-job hours and formula context"
        Next
    Next
End Sub

Function MathNumber(value)
    MathNumber = 0
    If IsNumeric(value) Then MathNumber = CDbl(value)
End Function

Sub RecordActionResult(workbookObj, currentRunId, scenarioName, stepName, actionResult, details)
    If Left(actionResult, 2) = "OK" Then
        RecordPass workbookObj, currentRunId, scenarioName, stepName, details
    Else
        RecordFail workbookObj, currentRunId, scenarioName, stepName, details & " | " & actionResult
    End If
End Sub

Sub AssertTrue(workbookObj, currentRunId, scenarioName, stepName, conditionValue, details)
    If conditionValue Then
        RecordPass workbookObj, currentRunId, scenarioName, stepName, details
    Else
        RecordFail workbookObj, currentRunId, scenarioName, stepName, details
    End If
End Sub

Sub AssertContains(workbookObj, currentRunId, scenarioName, stepName, actualText, expectedText, details)
    If InStr(1, actualText, expectedText, vbTextCompare) > 0 Then
        RecordPass workbookObj, currentRunId, scenarioName, stepName, details & " | " & actualText
    Else
        RecordFail workbookObj, currentRunId, scenarioName, stepName, details & " | Actual=" & actualText
    End If
End Sub

Sub RecordPass(workbookObj, currentRunId, scenarioName, stepName, details)
    On Error Resume Next
    workbookObj.Application.Run WorkbookMacro(workbookObj, "modInteractionLogging.AppendTestResult"), currentRunId, scenarioName, stepName, True, details, ""
    WriteLog "PASS | " & scenarioName & " | " & stepName & " | " & details
End Sub

Sub RecordFail(workbookObj, currentRunId, scenarioName, stepName, details)
    On Error Resume Next
    failureCount = failureCount + 1
    workbookObj.Application.Run WorkbookMacro(workbookObj, "modInteractionLogging.AppendTestResult"), currentRunId, scenarioName, stepName, False, details, ""
    WriteLog "FAIL | " & scenarioName & " | " & stepName & " | " & details
End Sub

Sub CaptureSheetWindowPng(excelApp, sheetObj, outputPath, scrollRow)
    Dim commandText
    On Error Resume Next
    sheetObj.Activate
    excelApp.ActiveWindow.ScrollRow = scrollRow
    excelApp.ActiveWindow.ScrollColumn = 1
    excelApp.ActiveWindow.Zoom = 90
    commandText = "powershell -NoProfile -ExecutionPolicy Bypass -File ""C:\dev\STLPackingPerformanceReport\scripts\capture-excel-window.ps1"" -WindowHandle " & CStr(excelApp.Hwnd) & " -OutputPath """ & outputPath & """"
    shell.Run commandText, 0, True
End Sub

Sub WriteDashboardLayoutAudit(sheetObj, outputPath, workbookObj, currentRunId)
    Dim ts
    Dim i
    Dim j
    Dim shpA
    Dim shpB
    Dim loA
    Dim loB
    Dim issueCount

    On Error Resume Next
    Set ts = fso.OpenTextFile(outputPath, 2, True)
    ts.WriteLine "IssueType,ShapeName,OtherShapeName,Details"

    For i = 1 To sheetObj.Shapes.Count
        Set shpA = sheetObj.Shapes(i)
        If IsAuditedDashboardShape(shpA) Then
            If shpA.Width < 72 Or shpA.Height < 20 Then
                ts.WriteLine "undersized," & shpA.Name & ",," & shpA.Width & "x" & shpA.Height
                issueCount = issueCount + 1
            End If
        End If
    Next

    For i = 1 To sheetObj.Shapes.Count
        Set shpA = sheetObj.Shapes(i)
        If IsAuditedDashboardShape(shpA) Then
            For j = i + 1 To sheetObj.Shapes.Count
                Set shpB = sheetObj.Shapes(j)
                If IsAuditedDashboardShape(shpB) Then
                    If RectanglesOverlap(shpA, shpB) Then
                        ts.WriteLine "shape_overlap," & shpA.Name & "," & shpB.Name & ",Dashboard shapes overlap"
                        issueCount = issueCount + 1
                    End If
                End If
            Next

            For Each loA In sheetObj.ListObjects
                If ShapeOverlapsListObject(shpA, loA) Then
                    ts.WriteLine "shape_table_overlap," & shpA.Name & "," & loA.Name & ",Shape overlaps dashboard table"
                    issueCount = issueCount + 1
                End If
            Next
        End If
    Next

    For i = 1 To sheetObj.ListObjects.Count
        Set loA = sheetObj.ListObjects(i)
        For j = i + 1 To sheetObj.ListObjects.Count
            Set loB = sheetObj.ListObjects(j)
            If TableRangesOverlap(loA, loB) Then
                ts.WriteLine "table_overlap," & loA.Name & "," & loB.Name & ",Dashboard tables overlap"
                issueCount = issueCount + 1
            End If
        Next
    Next

    ts.Close
    If issueCount = 0 Then
        RecordPass workbookObj, currentRunId, "Visual audit", "Dashboard layout audit", "No audited shape or table overlap issues detected"
    Else
        RecordFail workbookObj, currentRunId, "Visual audit", "Dashboard layout audit", CStr(issueCount) & " issues written to " & outputPath
    End If
End Sub

Function IsAuditedDashboardShape(shp)
    IsAuditedDashboardShape = (Left(LCase(CStr(shp.Name)), 7) = "ppr_btn")
End Function

Function ShapeOverlapsListObject(shp, lo)
    Dim rng
    Set rng = lo.Range
    ShapeOverlapsListObject = Not (shp.Left + shp.Width <= rng.Left Or _
                                   rng.Left + rng.Width <= shp.Left Or _
                                   shp.Top + shp.Height <= rng.Top Or _
                                   rng.Top + rng.Height <= shp.Top)
End Function

Function TableRangesOverlap(loA, loB)
    Dim rngA
    Dim rngB
    Set rngA = loA.Range
    Set rngB = loB.Range
    TableRangesOverlap = Not (rngA.Left + rngA.Width <= rngB.Left Or _
                              rngB.Left + rngB.Width <= rngA.Left Or _
                              rngA.Top + rngA.Height <= rngB.Top Or _
                              rngB.Top + rngB.Height <= rngA.Top)
End Function

Function RectanglesOverlap(shapeA, shapeB)
    RectanglesOverlap = Not (shapeA.Left + shapeA.Width <= shapeB.Left Or _
                             shapeB.Left + shapeB.Width <= shapeA.Left Or _
                             shapeA.Top + shapeA.Height <= shapeB.Top Or _
                             shapeB.Top + shapeB.Height <= shapeA.Top)
End Function

Function WorkbookMacro(workbookObj, macroName)
    WorkbookMacro = "'" & workbookObj.Name & "'!" & macroName
End Function

Sub EnsureFolder(folderPath)
    On Error Resume Next
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Sub WriteLog(messageText)
    Dim ts
    On Error Resume Next
    Set ts = fso.OpenTextFile(logPath, 8, True)
    ts.WriteLine Now & " | " & messageText
    ts.Close
End Sub