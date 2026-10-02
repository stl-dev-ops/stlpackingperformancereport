Attribute VB_Name = "modReportMain"
Option Explicit

Public Sub InitializeReportWorkbook()
    EnsureBaseSheets
    EnsureConfigurationSheet
    ApplyDefaultFilters
    SetupDashboardVisuals
    EnsureDataSheets
    FocusDashboardTop
    SetConfigValue NAME_VERSION, REPORT_VERSION
    UpdateDashboardStatus "Ready", RGB(255, 255, 255)
    AppendInteractionLog "InitializeReportWorkbook", "Success", "Workbook shell ready", DescribeCurrentFilters()
End Sub

Public Sub RefreshReport()
    Dim savedState As ExcelApplicationState
    Dim cn As Object
    Dim headers As Variant
    Dim data As Variant
    Dim workCenterHeaders As Variant
    Dim employeeHeaders As Variant
    Dim dashboardDetailHeaders As Variant
    Dim workCenterSummary As Variant
    Dim employeeSummary As Variant
    Dim dashboardDetailPreview As Variant
    Dim t0 As Double
    Dim durationMs As Double
    Dim rowCount As Long
    Dim failureMessage As String

    On Error GoTo CleanFail

    savedState = CaptureApplicationState()
    PrepareApplicationForRefresh
    RefreshDashboardSelectionCaptions
    ValidateConfiguredDateRange "Estimate Date", NAME_ESTIMATE_DATE_FROM, NAME_ESTIMATE_DATE_TO
    ValidateConfiguredDateRange "Delivery Date", NAME_DELIVERY_DATE_FROM, NAME_DELIVERY_DATE_TO
    ValidateConfiguredDateRange "Actual Work Date", NAME_ACTUAL_WORK_DATE_FROM, NAME_ACTUAL_WORK_DATE_TO
    AppendInteractionLog "RefreshReport", "Started", "Running packing refresh", DescribeCurrentFilters()

    t0 = Timer
    Set cn = OpenCermConnection()
    data = LoadPackingEstimatedVsActualData(cn, headers)
    durationMs = ElapsedMilliseconds(t0)
    rowCount = MatrixRowCount(data)
    workCenterSummary = BuildDimensionSummary(data, 8, "Work Center", workCenterHeaders)
    employeeSummary = BuildDimensionSummary(data, 9, "Employee", employeeHeaders)
    dashboardDetailHeaders = GetDashboardDetailHeaders()
    dashboardDetailPreview = BuildDashboardDetailPreview(data)

    WriteMatrixToTable EnsureWorksheet(SHEET_PACKING_DETAIL), TABLE_PACKING_DETAIL, "A1", headers, data
    ApplyPackingTableFormatting
    WriteDashboardDetailGrid EnsureWorksheet(SHEET_DASHBOARD), dashboardDetailHeaders, dashboardDetailPreview
    WriteMatrixToTable EnsureWorksheet(SHEET_DASHBOARD), TABLE_WORKCENTER_SUMMARY, "B31", workCenterHeaders, workCenterSummary
    WriteMatrixToTable EnsureWorksheet(SHEET_DASHBOARD), TABLE_EMPLOYEE_SUMMARY, "J31", employeeHeaders, employeeSummary
    ApplyDashboardDetailFormatting EnsureWorksheet(SHEET_DASHBOARD)
    ApplySummaryTableFormatting EnsureWorksheet(SHEET_DASHBOARD), TABLE_WORKCENTER_SUMMARY
    ApplySummaryTableFormatting EnsureWorksheet(SHEET_DASHBOARD), TABLE_EMPLOYEE_SUMMARY
    ApplySummaryJobComments TABLE_WORKCENTER_SUMMARY, data, 8
    ApplySummaryJobComments TABLE_EMPLOYEE_SUMMARY, data, 9
    UpdateDashboardMetrics data
    UpdateLastRefreshStamp
    AppendQueryLog "Success", rowCount, durationMs, DescribeCurrentFilters(), ""
    UpdateDashboardStatus "Refresh complete: " & CStr(rowCount) & " row(s)", RGB(255, 255, 255)
    AppendInteractionLog "RefreshReport", "Success", "Loaded " & CStr(rowCount) & " row(s)", DescribeCurrentFilters()

CleanExit:
    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
    End If
    RestoreApplicationState savedState
    Exit Sub

CleanFail:
    durationMs = ElapsedMilliseconds(t0)
    failureMessage = BuildFailureMessage(Err.Number, Err.Description)
    AppendQueryLog "Failure", 0, durationMs, DescribeCurrentFilters(), failureMessage
    UpdateDashboardStatus "Refresh failed: " & failureMessage, RGB(255, 199, 206)
    AppendInteractionLog "RefreshReport", "Failure", failureMessage, DescribeCurrentFilters()
    Resume CleanExit
End Sub

Public Sub EditEstimateDateRange()
    PromptForDateRange "Estimate Date", NAME_ESTIMATE_DATE_FROM, NAME_ESTIMATE_DATE_TO
End Sub

Public Sub EditDeliveryDateRange()
    PromptForDateRange "Delivery Date", NAME_DELIVERY_DATE_FROM, NAME_DELIVERY_DATE_TO
End Sub

Public Sub EditActualWorkDateRange()
    PromptForDateRange "Actual Work Date", NAME_ACTUAL_WORK_DATE_FROM, NAME_ACTUAL_WORK_DATE_TO
End Sub

Public Sub EditCustomerLike()
    SelectCustomerFilter
End Sub

Public Sub EditWorkCenterLike()
    SelectWorkCenterFilter
End Sub

Public Sub EditEmployeeLike()
    SelectEmployeeFilter
End Sub

Public Sub ClearAllFilters()
    ApplyDefaultFilters
    RefreshDashboardSelectionCaptions
    UpdateDashboardStatus "Defaults restored. Click Refresh to apply.", RGB(255, 255, 255)
    AppendInteractionLog "ClearAllFilters", "Success", "Default filters restored", DescribeCurrentFilters()
End Sub

Public Sub GoToConfigurationSheet()
    EnsureWorksheet(SHEET_CONFIGURATION).Activate
    AppendInteractionLog "Navigation", "Success", "Configuration", DescribeCurrentFilters()
End Sub

Public Sub GoToDashboardSheet()
    EnsureWorksheet(SHEET_DASHBOARD).Activate
    FocusDashboardTop
    AppendInteractionLog "Navigation", "Success", "Dashboard", DescribeCurrentFilters()
End Sub

Public Sub GoToDetailSheet()
    EnsureWorksheet(SHEET_PACKING_DETAIL).Activate
    AppendInteractionLog "Navigation", "Success", "Packing Detail", DescribeCurrentFilters()
End Sub

Public Sub GoToQueryLogSheet()
    EnsureWorksheet(SHEET_QUERY_LOG).Activate
    AppendInteractionLog "Navigation", "Success", "Query Log", DescribeCurrentFilters()
End Sub

Private Sub EnsureBaseSheets()
    EnsureWorksheet SHEET_DASHBOARD
    EnsureWorksheet SHEET_CONFIGURATION
    EnsureWorksheet SHEET_PACKING_DETAIL
    EnsureWorksheet SHEET_QUERY_LOG
    EnsureWorksheet SHEET_INTERACTION_LOG
    EnsureWorksheet SHEET_TEST_RESULTS
End Sub

Private Sub EnsureConfigurationSheet()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_CONFIGURATION)
    ws.Cells.Clear
    ws.Cells.Font.Name = "Bahnschrift"
    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 56

    ws.Range("A1:C1").Merge
    ws.Range("A1").Value = REPORT_TITLE & " Configuration"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A3:C3").Value = Array("Setting", "Value", "Notes")
    ws.Range("A3:C3").Font.Bold = True

    WriteConfigRow ws, 4, "Estimate Date From", NAME_ESTIMATE_DATE_FROM, "", "Optional inclusive lower bound for Estimate Date."
    WriteConfigRow ws, 5, "Estimate Date To", NAME_ESTIMATE_DATE_TO, "", "Optional inclusive upper bound for Estimate Date."
    WriteConfigRow ws, 6, "Delivery Date From", NAME_DELIVERY_DATE_FROM, "", "Optional inclusive lower bound for Delivery Date."
    WriteConfigRow ws, 7, "Delivery Date To", NAME_DELIVERY_DATE_TO, "", "Optional inclusive upper bound for Delivery Date."
    WriteConfigRow ws, 8, "Actual Work Date From", NAME_ACTUAL_WORK_DATE_FROM, LastWeekStartDate(), "Default inclusive lower bound for Actual Work Date (last week)."
    WriteConfigRow ws, 9, "Actual Work Date To", NAME_ACTUAL_WORK_DATE_TO, LastWeekEndDate(), "Default inclusive upper bound for Actual Work Date (last week)."
    WriteConfigRow ws, 10, "Customer Like", NAME_CUSTOMER_LIKE, "", "Optional SQL LIKE filter. Example: %Acme%"
    WriteConfigRow ws, 11, "Work Center Like", NAME_WORK_CENTER_LIKE, "Shipping,Shipping 2,Shipping 3", "Default work center subset for dashboard startup."
    WriteConfigRow ws, 12, "Employee Like", NAME_EMPLOYEE_LIKE, "", "Optional SQL LIKE filter."
    WriteConfigRow ws, 13, "Server Name", NAME_SERVER_NAME, "STL-SQL1\CRMDB", "Trusted connection SQL Server name."
    WriteConfigRow ws, 14, "Database Name", NAME_DATABASE_NAME, "sqlb00", "Initial catalog for the report query."
    WriteConfigRow ws, 15, "Connection Timeout", NAME_CONNECTION_TIMEOUT, 15, "Seconds."
    WriteConfigRow ws, 16, "Command Timeout", NAME_COMMAND_TIMEOUT, 120, "Seconds."
    WriteConfigRow ws, 17, "Workbook Version", NAME_VERSION, REPORT_VERSION, "Managed by workbook initialization."
    WriteConfigRow ws, 18, "Status", NAME_STATUS, "Ready", "Latest refresh status message."

    ws.Range("B4:B9").NumberFormat = "m/d/yyyy"
    ws.Range("A:C").VerticalAlignment = xlCenter
End Sub

Private Sub EnsureDataSheets()
    Dim detailHeaders As Variant
    Dim logHeaders As Variant
    Dim interactionHeaders As Variant
    Dim testHeaders As Variant
    Dim summaryHeaders As Variant
    Dim dashboardDetailHeaders As Variant

    detailHeaders = ArrayFromCsv("Job ID,Estimate Date,Delivery Date,Customer,Job Description,Order Quantity,Actual Work Date,Work Center,Employee,Estimated Packing Minutes,Actual Packing Minutes,Total Actual Packing Minutes,Share of Job Actual Time,Packing Minutes Variance,Variance %,Packing Status")
    logHeaders = ArrayFromCsv("Run At,Status,Rows,Duration Ms,Server,Database,Details,Filters")
    interactionHeaders = NormalizeHeaderArray(Array("Timestamp", "WindowsUsername", "EventName", "Status", "Details", "SelectionContext", "WorkbookVersion"))
    testHeaders = NormalizeHeaderArray(Array("RunId", "Timestamp", "Scenario", "StepName", "Status", "Details", "ArtifactPath"))
    summaryHeaders = ArrayFromCsv("Dimension,Jobs,Est Hrs,Act Hrs,Var Hrs,Variance %")
    dashboardDetailHeaders = GetDashboardDetailHeaders()

    WriteMatrixToTable EnsureWorksheet(SHEET_PACKING_DETAIL), TABLE_PACKING_DETAIL, "A1", detailHeaders, Empty
    WriteMatrixToTable EnsureWorksheet(SHEET_QUERY_LOG), TABLE_QUERY_LOG, "A1", logHeaders, Empty
    WriteMatrixToTable EnsureWorksheet(SHEET_INTERACTION_LOG), TABLE_INTERACTION_LOG, "A1", interactionHeaders, Empty
    WriteMatrixToTable EnsureWorksheet(SHEET_TEST_RESULTS), TABLE_TEST_RESULTS, "A1", testHeaders, Empty
    WriteDashboardDetailGrid EnsureWorksheet(SHEET_DASHBOARD), dashboardDetailHeaders, Empty
    WriteMatrixToTable EnsureWorksheet(SHEET_DASHBOARD), TABLE_WORKCENTER_SUMMARY, "B31", summaryHeaders, Empty
    WriteMatrixToTable EnsureWorksheet(SHEET_DASHBOARD), TABLE_EMPLOYEE_SUMMARY, "J31", summaryHeaders, Empty
End Sub

Public Sub FocusDashboardTop()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    ws.Activate
    ws.Range("B2").Select
    On Error Resume Next
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    On Error GoTo 0
End Sub

Private Sub ApplyDefaultFilters()
    SetConfigValue NAME_ESTIMATE_DATE_FROM, Empty
    SetConfigValue NAME_ESTIMATE_DATE_TO, Empty
    SetConfigValue NAME_DELIVERY_DATE_FROM, Empty
    SetConfigValue NAME_DELIVERY_DATE_TO, Empty
    SetConfigValue NAME_ACTUAL_WORK_DATE_FROM, LastWeekStartDate()
    SetConfigValue NAME_ACTUAL_WORK_DATE_TO, LastWeekEndDate()
    SetConfigValue NAME_CUSTOMER_LIKE, ""
    SetConfigValue NAME_WORK_CENTER_LIKE, "Shipping,Shipping 2,Shipping 3"
    SetConfigValue NAME_EMPLOYEE_LIKE, ""
End Sub

Private Function GetDashboardDetailHeaders() As Variant
    GetDashboardDetailHeaders = ArrayFromCsv("Jobs,Customers,Estimated Hours,Actual Hours,Variance Hours,Over Target Jobs,Variance %,Packing Status,Job Description,Order Quantity,Actual Work Date,Work Center,Employee")
End Function

Private Function BuildDashboardDetailPreview(ByVal dataMatrix As Variant) As Variant
    Const MAX_PREVIEW_ROWS As Long = 12
    Dim previewRows As Long
    Dim rowIndex As Long
    Dim result() As Variant

    previewRows = MatrixRowCount(dataMatrix)
    If previewRows > MAX_PREVIEW_ROWS Then previewRows = MAX_PREVIEW_ROWS
    If previewRows <= 0 Then Exit Function

    ReDim result(1 To previewRows, 1 To 13)
    For rowIndex = 1 To previewRows
        result(rowIndex, 1) = dataMatrix(rowIndex, 1)
        result(rowIndex, 2) = dataMatrix(rowIndex, 4)
        result(rowIndex, 3) = NzNumber(dataMatrix(rowIndex, 10)) / 60#
        result(rowIndex, 4) = NzNumber(dataMatrix(rowIndex, 11)) / 60#
        result(rowIndex, 5) = NzNumber(dataMatrix(rowIndex, 14)) / 60#
        If NzNumber(dataMatrix(rowIndex, 11)) > NzNumber(dataMatrix(rowIndex, 10)) Then
            result(rowIndex, 6) = 1
        Else
            result(rowIndex, 6) = 0
        End If
        result(rowIndex, 7) = PackingPercentage(NzNumber(dataMatrix(rowIndex, 11)), NzNumber(dataMatrix(rowIndex, 10)))
        result(rowIndex, 8) = dataMatrix(rowIndex, 16)
        result(rowIndex, 9) = dataMatrix(rowIndex, 5)
        result(rowIndex, 10) = dataMatrix(rowIndex, 6)
        result(rowIndex, 11) = dataMatrix(rowIndex, 7)
        result(rowIndex, 12) = dataMatrix(rowIndex, 8)
        result(rowIndex, 13) = dataMatrix(rowIndex, 9)
    Next rowIndex

    BuildDashboardDetailPreview = result
End Function

Private Sub WriteConfigRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal labelText As String, ByVal rangeName As String, ByVal defaultValue As Variant, ByVal noteText As String)
    ws.Cells(rowIndex, 1).Value = labelText
    ws.Cells(rowIndex, 2).Value = defaultValue
    ws.Cells(rowIndex, 3).Value = noteText
    DefineWorkbookName rangeName, ws.Cells(rowIndex, 2)
End Sub

Private Sub DefineWorkbookName(ByVal rangeName As String, ByVal targetCell As Range)
    On Error Resume Next
    ThisWorkbook.Names(rangeName).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=rangeName, RefersTo:="=" & targetCell.Address(True, True, xlA1, True)
End Sub

Private Sub AppendQueryLog(ByVal statusText As String, ByVal rowCount As Long, ByVal durationMs As Double, ByVal filterText As String, ByVal detailText As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim newRow As ListRow

    Set ws = EnsureWorksheet(SHEET_QUERY_LOG)
    Set lo = ws.ListObjects(TABLE_QUERY_LOG)
    Set newRow = lo.ListRows.Add

    With newRow.Range
        .Cells(1, 1).Value = Now
        .Cells(1, 2).Value = statusText
        .Cells(1, 3).Value = rowCount
        .Cells(1, 4).Value = durationMs
        .Cells(1, 5).Value = ServerName()
        .Cells(1, 6).Value = DatabaseName()
        .Cells(1, 7).Value = detailText
        .Cells(1, 8).Value = filterText
    End With

    lo.ListColumns("Run At").DataBodyRange.NumberFormat = "m/d/yyyy h:mm AM/PM"
    lo.ListColumns("Duration Ms").DataBodyRange.NumberFormat = "#,##0"
    lo.Range.EntireColumn.AutoFit
End Sub

Private Function BuildDimensionSummary(ByVal dataMatrix As Variant, ByVal keyColumnIndex As Long, ByVal headerText As String, ByRef headers As Variant) As Variant
    Dim stats As Object
    Dim rowIndex As Long
    Dim dimensionKey As String
    Dim detail As Object
    Dim keys() As String
    Dim outputRow As Long
    Dim result() As Variant
    Dim keyIndex As Long

    headers = ArrayFromCsv(headerText & ",Jobs,Est Hrs,Act Hrs,Var Hrs,Variance %")
    Set stats = CreateObject("Scripting.Dictionary")

    For rowIndex = 1 To MatrixRowCount(dataMatrix)
        dimensionKey = NzText(dataMatrix(rowIndex, keyColumnIndex), "(Blank)")
        If Not stats.Exists(dimensionKey) Then
            Set detail = CreateObject("Scripting.Dictionary")
            Set detail("Jobs") = CreateObject("Scripting.Dictionary")
            detail("Estimated") = 0#
            detail("Actual") = 0#
            detail("Variance") = 0#
            Set stats(dimensionKey) = detail
        End If

        Set detail = stats(dimensionKey)
        If Len(NzText(dataMatrix(rowIndex, 1))) > 0 Then detail("Jobs")(NzText(dataMatrix(rowIndex, 1))) = True
        detail("Estimated") = NzNumber(detail("Estimated")) + (NzNumber(dataMatrix(rowIndex, 10)) / 60#)
        detail("Actual") = NzNumber(detail("Actual")) + (NzNumber(dataMatrix(rowIndex, 11)) / 60#)
        detail("Variance") = NzNumber(detail("Variance")) + (NzNumber(dataMatrix(rowIndex, 14)) / 60#)
    Next rowIndex

    If stats.Count = 0 Then
        BuildDimensionSummary = Empty
        Exit Function
    End If

    keys = SortSummaryKeysByActualHours(stats)
    ReDim result(1 To UBound(keys) - LBound(keys) + 1, 1 To 6)
    outputRow = 0
    For keyIndex = LBound(keys) To UBound(keys)
        outputRow = outputRow + 1
        Set detail = stats(keys(keyIndex))
        result(outputRow, 1) = keys(keyIndex)
        result(outputRow, 2) = detail("Jobs").Count
        result(outputRow, 3) = NzNumber(detail("Estimated"))
        result(outputRow, 4) = NzNumber(detail("Actual"))
        result(outputRow, 5) = NzNumber(detail("Actual")) - NzNumber(detail("Estimated"))
        result(outputRow, 6) = PackingPercentage(NzNumber(detail("Actual")), NzNumber(detail("Estimated")))
    Next keyIndex

    BuildDimensionSummary = result
End Function

Public Function PackingPercentage(ByVal actualMinutes As Double, ByVal estimatedMinutes As Double) As Variant
    If estimatedMinutes > 0 Then
        PackingPercentage = (actualMinutes - estimatedMinutes) / estimatedMinutes
    Else
        PackingPercentage = "N/A"
    End If
End Function

Private Sub ApplySummaryJobComments(ByVal tableName As String, ByVal dataMatrix As Variant, ByVal dimensionColumn As Long)
    Dim lo As ListObject
    Dim summaryRow As Long
    Dim dataRow As Long
    Dim columnIndex As Long
    Dim groupName As String
    Dim jobId As String
    Dim jobs As Object
    Dim jobKey As Variant
    Dim jobHours As Variant
    Dim auditText As String
    Dim baseText As String
    Dim valueText As String
    Dim formulas As Variant

    Set lo = EnsureWorksheet(SHEET_DASHBOARD).ListObjects(tableName)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    ClearRangeComments lo.DataBodyRange
    formulas = Array("Jobs included in the active filters.", "Unique Job ID count.", _
        "Sum of allocated estimated hours.", "Sum of actual hours.", _
        "Actual hours - allocated estimated hours.", _
        "(Actual hours - allocated estimated hours)" & vbLf & _
        "         / allocated estimated hours." & vbLf & _
        "Negative = under; positive = over; 0% = on estimate." & vbLf & _
        "N/A = no positive estimate.")

    For summaryRow = 1 To lo.ListRows.Count
        groupName = NzText(lo.DataBodyRange.Cells(summaryRow, 1).Value)
        Set jobs = CreateObject("Scripting.Dictionary")
        For dataRow = 1 To MatrixRowCount(dataMatrix)
            If NzText(dataMatrix(dataRow, dimensionColumn), "(Blank)") = groupName Then
                jobId = NzText(dataMatrix(dataRow, 1), "(No Job ID)")
                If jobs.Exists(jobId) Then
                    jobHours = jobs(jobId)
                Else
                    jobHours = Array(0#, 0#)
                End If
                jobHours(0) = jobHours(0) + NzNumber(dataMatrix(dataRow, 10)) / 60#
                jobHours(1) = jobHours(1) + NzNumber(dataMatrix(dataRow, 11)) / 60#
                jobs(jobId) = jobHours
            End If
        Next dataRow
        auditText = PackingAuditLine("Job ID", "Est Hrs", "Act Hrs", "Var Hrs", "Var %")
        For Each jobKey In jobs.Keys
            jobHours = jobs(jobKey)
            valueText = "N/A"
            If jobHours(0) > 0 Then valueText = Format$(PackingPercentage(jobHours(1), jobHours(0)), "+0.00%;-0.00%;0.00%")
            auditText = auditText & vbLf & PackingAuditLine(CStr(jobKey), _
                Format$(jobHours(0), "0.00"), Format$(jobHours(1), "0.00"), _
                Format$(jobHours(1) - jobHours(0), "0.00"), valueText)
        Next jobKey
        For columnIndex = 1 To lo.ListColumns.Count
            valueText = NzText(lo.DataBodyRange.Cells(summaryRow, columnIndex).Value)
            If columnIndex >= 3 And columnIndex <= 5 Then valueText = Format$(lo.DataBodyRange.Cells(summaryRow, columnIndex).Value, "0.00")
            If columnIndex = 6 And IsNumeric(lo.DataBodyRange.Cells(summaryRow, columnIndex).Value) Then valueText = Format$(lo.DataBodyRange.Cells(summaryRow, columnIndex).Value, "+0.00%;-0.00%;0.00%")
            baseText = lo.HeaderRowRange.Cells(1, columnIndex).Value & vbLf & groupName & vbLf & _
                "Formula: " & formulas(columnIndex - 1) & vbLf & "Displayed value: " & valueText & vbLf & _
                "Estimates allocated by share of job time," & vbLf & _
                "not independent employee budgets." & vbLf & vbLf & _
                "Jobs included (active filters):" & vbLf & auditText
            SetCellComment lo.DataBodyRange.Cells(summaryRow, columnIndex), baseText
            With lo.DataBodyRange.Cells(summaryRow, columnIndex).Comment.Shape.TextFrame
                .Characters.Font.Name = "Courier New"
                .AutoSize = True
            End With
        Next columnIndex
    Next summaryRow
End Sub

Private Function PackingAuditLine(ByVal jobId As String, ByVal estimated As String, ByVal actual As String, ByVal variance As String, ByVal percentage As String) As String
    PackingAuditLine = jobId & Space$(Application.Max(1, 14 - Len(jobId))) & _
        Space$(Application.Max(0, 10 - Len(estimated))) & estimated & " " & _
        Space$(Application.Max(0, 10 - Len(actual))) & actual & " " & _
        Space$(Application.Max(0, 10 - Len(variance))) & variance & " " & _
        Space$(Application.Max(0, 10 - Len(percentage))) & percentage
End Function

Private Function SortSummaryKeysByActualHours(ByVal stats As Object) As String()
    Dim keys() As String
    Dim index As Long
    Dim innerIndex As Long
    Dim temp As String
    Dim someKey As Variant

    ReDim keys(0 To stats.Count - 1)
    index = 0
    For Each someKey In stats.Keys
        keys(index) = CStr(someKey)
        index = index + 1
    Next someKey

    For index = LBound(keys) To UBound(keys) - 1
        For innerIndex = index + 1 To UBound(keys)
            If NzNumber(stats(keys(innerIndex))("Actual")) > NzNumber(stats(keys(index))("Actual")) Then
                temp = keys(index)
                keys(index) = keys(innerIndex)
                keys(innerIndex) = temp
            End If
        Next innerIndex
    Next index

    SortSummaryKeysByActualHours = keys
End Function

Private Sub ValidateConfiguredDateRange(ByVal labelText As String, ByVal fromName As String, ByVal toName As String)
    Dim fromValue As Variant
    Dim toValue As Variant

    fromValue = GetConfigDate(fromName)
    toValue = GetConfigDate(toName)
    If Not IsEmpty(fromValue) And Not IsEmpty(toValue) Then
        If CDate(toValue) < CDate(fromValue) Then
            Err.Raise vbObjectError + 601, "ValidateConfiguredDateRange", labelText & " end date cannot be earlier than start date."
        End If
    End If
End Sub

Private Sub PromptForDateRange(ByVal labelText As String, ByVal fromName As String, ByVal toName As String)
    Dim fromInput As String
    Dim toInput As String
    Dim currentFrom As Variant
    Dim currentTo As Variant

    currentFrom = GetConfigDate(fromName)
    currentTo = GetConfigDate(toName)

    fromInput = InputBox(labelText & " start date (m/d/yyyy). Leave blank to clear.", REPORT_TITLE, IIf(IsEmpty(currentFrom), "", Format$(CDate(currentFrom), "m/d/yyyy")))
    If StrPtr(fromInput) = 0 Then Exit Sub
    toInput = InputBox(labelText & " end date (m/d/yyyy). Leave blank to clear.", REPORT_TITLE, IIf(IsEmpty(currentTo), "", Format$(CDate(currentTo), "m/d/yyyy")))
    If StrPtr(toInput) = 0 Then Exit Sub

    If Len(Trim$(fromInput)) = 0 Then
        SetConfigValue fromName, Empty
    ElseIf IsDate(fromInput) Then
        SetConfigValue fromName, CDate(fromInput)
    Else
        Err.Raise vbObjectError + 602, "PromptForDateRange", labelText & " start date is not a valid date."
    End If

    If Len(Trim$(toInput)) = 0 Then
        SetConfigValue toName, Empty
    ElseIf IsDate(toInput) Then
        SetConfigValue toName, CDate(toInput)
    Else
        Err.Raise vbObjectError + 603, "PromptForDateRange", labelText & " end date is not a valid date."
    End If

    RefreshDashboardSelectionCaptions
    UpdateDashboardStatus labelText & " updated. Click Refresh to apply.", RGB(255, 255, 255)
    AppendInteractionLog "FilterUpdate", "Success", labelText & " updated", DescribeCurrentFilters()
End Sub

Private Sub PromptForTextFilter(ByVal labelText As String, ByVal rangeName As String, ByVal placeholder As String)
    Dim currentValue As String
    Dim newValue As String

    currentValue = GetConfigText(rangeName)
    newValue = InputBox(labelText & " filter. Leave blank to clear. Example: " & placeholder, REPORT_TITLE, currentValue)
    If StrPtr(newValue) = 0 Then Exit Sub

    SetConfigValue rangeName, Trim$(newValue)
    RefreshDashboardSelectionCaptions
    UpdateDashboardStatus labelText & " updated. Click Refresh to apply.", RGB(255, 255, 255)
    AppendInteractionLog "FilterUpdate", "Success", labelText & " updated", DescribeCurrentFilters()
End Sub

Public Function DescribeCurrentFilters() As String
    DescribeCurrentFilters = Join(Array( _
        "Estimate=" & DescribeFilterPair(GetConfigDate(NAME_ESTIMATE_DATE_FROM), GetConfigDate(NAME_ESTIMATE_DATE_TO)), _
        "Delivery=" & DescribeFilterPair(GetConfigDate(NAME_DELIVERY_DATE_FROM), GetConfigDate(NAME_DELIVERY_DATE_TO)), _
        "ActualWork=" & DescribeFilterPair(GetConfigDate(NAME_ACTUAL_WORK_DATE_FROM), GetConfigDate(NAME_ACTUAL_WORK_DATE_TO)), _
        "Customer=" & DefaultText(GetConfigText(NAME_CUSTOMER_LIKE), "NULL"), _
        "WorkCenter=" & DefaultText(GetConfigText(NAME_WORK_CENTER_LIKE), "NULL"), _
        "Employee=" & DefaultText(GetConfigText(NAME_EMPLOYEE_LIKE), "NULL") _
    ), " | ")
End Function

Private Function DescribeFilterPair(ByVal fromValue As Variant, ByVal toValue As Variant) As String
    DescribeFilterPair = DefaultVariantDateText(fromValue) & ".." & DefaultVariantDateText(toValue)
End Function

Private Function DefaultVariantDateText(ByVal someDate As Variant) As String
    If IsEmpty(someDate) Then
        DefaultVariantDateText = "NULL"
    Else
        DefaultVariantDateText = Format$(CDate(someDate), "yyyy-mm-dd")
    End If
End Function

Private Function DefaultText(ByVal valueText As String, ByVal fallbackText As String) As String
    If Len(Trim$(valueText)) = 0 Then
        DefaultText = fallbackText
    Else
        DefaultText = valueText
    End If
End Function

Private Function ArrayFromCsv(ByVal csvText As String) As Variant
    Dim rawItems As Variant
    Dim result() As Variant
    Dim index As Long

    rawItems = Split(csvText, ",")
    ReDim result(1 To UBound(rawItems) - LBound(rawItems) + 1)
    For index = LBound(rawItems) To UBound(rawItems)
        result(index - LBound(rawItems) + 1) = Trim$(CStr(rawItems(index)))
    Next index
    ArrayFromCsv = result
End Function

Private Function BuildFailureMessage(ByVal failureNumber As Long, ByVal failureDescription As String) As String
    Dim messageText As String

    messageText = SanitizedErrorMessage(failureDescription)
    If Len(Trim$(messageText)) = 0 Then
        BuildFailureMessage = "VBA error " & CStr(failureNumber)
    Else
        BuildFailureMessage = messageText
    End If
End Function
