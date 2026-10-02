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

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

artifactsRoot = "C:\dev\STLPackingPerformanceReport\artifacts"
exportRoot = artifactsRoot & "\e2e"
logPath = exportRoot & "\run-e2e.log"
layoutAuditPath = exportRoot & "\dashboard-layout-audit.csv"
builtWorkbookPath = "C:\dev\STLPackingPerformanceReport\PackingEstimatedVsActualReport.xlsm"
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
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Open(builtWorkbookPath)
Set wsDash = wb.Worksheets("Dashboard")
xl.WindowState = -4137

runId = InvokeAction(xl, wb, "new-run-id", "e2e", "")
WriteLog "Run ID: " & runId

RecordPass wb, runId, "Workbook opens", "Open workbook and initialize", "Workbook opened for automation"
AssertContains wb, runId, "Workbook opens", "Dashboard sheet active on open", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCurrentSheetName"))), "Dashboard", "Workbook should open on Dashboard"
RecordActionResult wb, runId, "Initialize", "InitializeReportWorkbook", InvokeAction(xl, wb, "initialize", "", ""), "Workbook initializer completed"
AssertContains wb, runId, "Initialize", "Dashboard sheet active after initialize", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCurrentSheetName"))), "Dashboard", "Initializer should leave workbook on Dashboard"

AssertTrue wb, runId, "Dashboard controls", "Estimate-from button exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "ppr_btnEstimateDateFrom")), "Estimate-from button should render"
AssertTrue wb, runId, "Dashboard controls", "Refresh button exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "ppr_btnRefresh")), "Refresh button should render"
AssertTrue wb, runId, "Dashboard controls", "KPI comment exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCommentExists"), "Dashboard", "B17")), "KPI tooltip comment should exist"
AssertTrue wb, runId, "Dashboard controls", "Customer button exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "ppr_btnCustomerFilter")), "Customer picker button should render"

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

RecordActionResult wb, runId, "Picker workflow", "Clear filters for subset workflow", InvokeAction(xl, wb, "clear-filters", "", ""), "Reset filter state"
AssertTrue wb, runId, "Picker workflow", "Work center filter cleared", Len(CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestConfigText"), "cfgWorkCenterLike"))) = 0, "Work center filter should be cleared"
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
AssertTrue wb, runId, "Picker workflow", "Distinct jobs KPI updates for Shipping subset", CDbl(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestKpiNumericValue"), "B17")) > 0, "Distinct jobs KPI should be non-zero for Shipping subset"
AssertTrue wb, runId, "Picker workflow", "Actual hours KPI updates for Shipping subset", CDbl(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestKpiNumericValue"), "H17")) > 0, "Actual hours KPI should be non-zero for Shipping subset"
AssertTrue wb, runId, "Picker workflow", "Work center summary has rows", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Dashboard", "tblWorkCenterSummary")) > 0, "Work center summary should populate for Shipping subset"
AssertTrue wb, runId, "Picker workflow", "Employee summary has rows", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Dashboard", "tblEmployeeSummary")) > 0, "Employee summary should populate for Shipping subset"

RecordActionResult wb, runId, "Reset workflow", "Clear filters after subset workflow", InvokeAction(xl, wb, "clear-filters", "", ""), "Returned dashboard to unfiltered state"
AssertContains wb, runId, "Reset workflow", "Active filter caption reset", CStr(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCellText"), "Dashboard", "B26")), "Work Center: Any", "Active filter panel should show cleared work center filter"
RecordActionResult wb, runId, "Reset workflow", "Return to dashboard before save", InvokeAction(xl, wb, "goto-dashboard", "", ""), "Ensure workbook ends on dashboard"

wb.Save
CaptureSheetWindowPng xl, wb.Worksheets("Dashboard"), exportRoot & "\dashboard.png"
CaptureSheetWindowPng xl, wb.Worksheets("Packing Detail"), exportRoot & "\packing-detail.png"
CaptureSheetWindowPng xl, wb.Worksheets("Query Log"), exportRoot & "\query-log.png"
WriteDashboardLayoutAudit wb.Worksheets("Dashboard"), layoutAuditPath, wb, runId

AssertTrue wb, runId, "Logging", "Interaction log populated", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Interaction Log", "tblInteractionLog")) > 0, "Interaction log should have entries"
AssertTrue wb, runId, "Logging", "Query log populated", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Query Log", "tblQueryLog")) > 0, "Query log should have entries"
AssertTrue wb, runId, "Logging", "Test results populated", CLng(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestTableRowCount"), "Test Results", "tblTestResults")) > 0, "Test results should have entries"

wb.Save
wb.Close False
xl.Quit

WriteLog "E2E completed"
WScript.Echo "OK:" & exportRoot

Function InvokeAction(excelApp, workbookObj, actionName, arg1, arg2)
    On Error Resume Next
    Err.Clear
    InvokeAction = CStr(excelApp.Run(WorkbookMacro(workbookObj, "modAutomationSupport.AutomationInvoke"), actionName, arg1, arg2))
    If Err.Number <> 0 Then InvokeAction = "ERR:" & Err.Number & ":" & Err.Description
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
    workbookObj.Application.Run WorkbookMacro(workbookObj, "modInteractionLogging.AppendTestResult"), currentRunId, scenarioName, stepName, False, details, ""
    WriteLog "FAIL | " & scenarioName & " | " & stepName & " | " & details
End Sub

Sub CaptureSheetWindowPng(excelApp, sheetObj, outputPath)
    Dim commandText
    On Error Resume Next
    sheetObj.Activate
    excelApp.ActiveWindow.ScrollRow = 1
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
    Dim issueCount

    On Error Resume Next
    Set ts = fso.OpenTextFile(outputPath, 2, True)
    ts.WriteLine "IssueType,ShapeName,OtherShapeName,Details"

    For i = 1 To sheetObj.Shapes.Count
        Set shpA = sheetObj.Shapes(i)
        If Left(LCase(shpA.Name), 7) = "ppr_btn" Then
            If shpA.Width < 72 Or shpA.Height < 20 Then
                ts.WriteLine "undersized," & shpA.Name & ",," & shpA.Width & "x" & shpA.Height
                issueCount = issueCount + 1
            End If
        End If
    Next

    For i = 1 To sheetObj.Shapes.Count
        Set shpA = sheetObj.Shapes(i)
        If Left(LCase(shpA.Name), 7) = "ppr_btn" Then
            For j = i + 1 To sheetObj.Shapes.Count
                Set shpB = sheetObj.Shapes(j)
                If Left(LCase(shpB.Name), 7) = "ppr_btn" Then
                    If RectanglesOverlap(shpA, shpB) Then
                        ts.WriteLine "overlap," & shpA.Name & "," & shpB.Name & ",Buttons overlap"
                        issueCount = issueCount + 1
                    End If
                End If
            Next
        End If
    Next

    ts.Close
    If issueCount = 0 Then
        RecordPass workbookObj, currentRunId, "Visual audit", "Dashboard layout audit", "No button overlap issues detected"
    Else
        RecordFail workbookObj, currentRunId, "Visual audit", "Dashboard layout audit", CStr(issueCount) & " issues written to " & outputPath
    End If
End Sub

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