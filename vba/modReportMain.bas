Attribute VB_Name = "modReportMain"
Option Explicit

Public Sub InitializeReportWorkbook()
    EnsureBaseSheets
    EnsureConfigurationSheet
    SetupDashboardVisuals
    EnsureDataSheets
    SetConfigValue NAME_VERSION, REPORT_VERSION
    UpdateDashboardStatus "Ready", RGB(255, 255, 255)
End Sub

Public Sub RefreshReport()
    Dim savedState As ExcelApplicationState
    Dim cn As Object
    Dim headers As Variant
    Dim data As Variant
    Dim t0 As Double
    Dim durationMs As Double
    Dim rowCount As Long

    On Error GoTo CleanFail

    savedState = CaptureApplicationState()
    PrepareApplicationForRefresh
    RefreshDashboardSelectionCaptions

    t0 = Timer
    Set cn = OpenCermConnection()
    data = LoadPackingEstimatedVsActualData(cn, headers)
    durationMs = ElapsedMilliseconds(t0)
    rowCount = MatrixRowCount(data)

    WriteMatrixToTable EnsureWorksheet(SHEET_PACKING_DETAIL), TABLE_PACKING_DETAIL, "A1", headers, data
    ApplyPackingTableFormatting
    UpdateDashboardMetrics data
    UpdateLastRefreshStamp
    AppendQueryLog "Success", rowCount, durationMs, DescribeCurrentFilters(), ""
    UpdateDashboardStatus "Refresh complete: " & CStr(rowCount) & " row(s)", RGB(255, 255, 255)

CleanExit:
    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
    End If
    RestoreApplicationState savedState
    Exit Sub

CleanFail:
    durationMs = ElapsedMilliseconds(t0)
    AppendQueryLog "Failure", 0, durationMs, DescribeCurrentFilters(), SanitizedErrorMessage(Err.Description)
    UpdateDashboardStatus "Refresh failed: " & SanitizedErrorMessage(Err.Description), RGB(255, 199, 206)
    Resume CleanExit
End Sub

Public Sub GoToConfigurationSheet()
    EnsureWorksheet(SHEET_CONFIGURATION).Activate
End Sub

Private Sub EnsureBaseSheets()
    EnsureWorksheet SHEET_DASHBOARD
    EnsureWorksheet SHEET_CONFIGURATION
    EnsureWorksheet SHEET_PACKING_DETAIL
    EnsureWorksheet SHEET_QUERY_LOG
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
    WriteConfigRow ws, 8, "Actual Work Date From", NAME_ACTUAL_WORK_DATE_FROM, "", "Optional inclusive lower bound for Actual Work Date."
    WriteConfigRow ws, 9, "Actual Work Date To", NAME_ACTUAL_WORK_DATE_TO, "", "Optional inclusive upper bound for Actual Work Date."
    WriteConfigRow ws, 10, "Customer Like", NAME_CUSTOMER_LIKE, "", "Optional SQL LIKE filter. Example: %Acme%"
    WriteConfigRow ws, 11, "Work Center Like", NAME_WORK_CENTER_LIKE, "", "Optional SQL LIKE filter."
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

    detailHeaders = ArrayFromCsv("Job ID,Estimate Date,Delivery Date,Customer,Job Description,Order Quantity,Actual Work Date,Work Center,Employee,Estimated Packing Minutes,Actual Packing Minutes,Total Actual Packing Minutes,Share of Job Actual Time,Packing Minutes Variance,Packing Time Ratio,Packing Status")
    logHeaders = ArrayFromCsv("Run At,Status,Rows,Duration Ms,Server,Database,Details,Filters")

    WriteMatrixToTable EnsureWorksheet(SHEET_PACKING_DETAIL), TABLE_PACKING_DETAIL, "A1", detailHeaders, Empty
    WriteMatrixToTable EnsureWorksheet(SHEET_QUERY_LOG), TABLE_QUERY_LOG, "A1", logHeaders, Empty
End Sub

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

Private Function DescribeCurrentFilters() As String
    DescribeCurrentFilters = Join(Array( _
        "Estimate=" & DescribeFilterPair(GetConfigDate(NAME_ESTIMATE_DATE_FROM), GetConfigDate(NAME_ESTIMATE_DATE_TO)), _
        "Delivery=" & DescribeFilterPair(GetConfigDate(NAME_DELIVERY_DATE_FROM), GetConfigDate(NAME_DELIVERY_DATE_TO)), _
        "ActualWork=" & DescribeFilterPair(GetConfigDate(NAME_ACTUAL_WORK_DATE_FROM), GetConfigDate(NAME_ACTUAL_WORK_DATE_TO)), _
        "CustomerLike=" & DefaultText(GetConfigText(NAME_CUSTOMER_LIKE), "NULL"), _
        "WorkCenterLike=" & DefaultText(GetConfigText(NAME_WORK_CENTER_LIKE), "NULL"), _
        "EmployeeLike=" & DefaultText(GetConfigText(NAME_EMPLOYEE_LIKE), "NULL") _
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
