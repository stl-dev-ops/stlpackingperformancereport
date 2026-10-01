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
RecordActionResult wb, runId, "Initialize", "InitializeReportWorkbook", InvokeAction(xl, wb, "initialize", "", ""), "Workbook initializer completed"

AssertTrue wb, runId, "Dashboard controls", "Estimate button exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "ppr_btnEstimateDates")), "Estimate button should render"
AssertTrue wb, runId, "Dashboard controls", "Refresh button exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestShapeExists"), "Dashboard", "ppr_btnRefresh")), "Refresh button should render"
AssertTrue wb, runId, "Dashboard controls", "KPI comment exists", CBool(xl.Run(WorkbookMacro(wb, "modAutomationSupport.TestCommentExists"), "Dashboard", "B17")), "KPI tooltip comment should exist"

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