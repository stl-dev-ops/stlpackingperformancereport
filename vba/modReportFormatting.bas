Attribute VB_Name = "modReportFormatting"
Option Explicit

Public Sub SetShapeButtonText(ByVal shp As Shape, ByVal caption As String)
    On Error Resume Next
    shp.TextFrame2.TextRange.Text = caption
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Name = "Bahnschrift SemiBold"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_TEXT_DARK
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Color = COLOR_TEXT_DARK
    shp.TextFrame.Characters.Font.Name = "Bahnschrift SemiBold"
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    On Error GoTo 0
End Sub

Public Function ShapeExists(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    On Error Resume Next
    ShapeExists = Len(ws.Shapes(shapeName).Name) > 0
    On Error GoTo 0
End Function

Public Sub DeleteShapesWithPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    Dim index As Long

    For index = ws.Shapes.Count To 1 Step -1
        If LCase$(Left$(ws.Shapes(index).Name, Len(prefix))) = LCase$(prefix) Then
            ws.Shapes(index).Delete
        End If
    Next index
End Sub

Public Sub AddButtonOnRange(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, ByVal macroName As String, ByVal anchor As Range, ByVal fillColor As Long)
    Dim shp As Shape

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left + 1, anchor.Top + 1, Application.Max(72, anchor.Width - 2), Application.Max(24, anchor.Height - 2))
    With shp
        .Name = shapeName
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = COLOR_BORDER
        .Line.Weight = 1
    End With
    SetShapeButtonText shp, caption
End Sub

Public Sub SetCellComment(ByVal targetCell As Range, ByVal commentText As String)
    On Error Resume Next
    If Not targetCell.Comment Is Nothing Then targetCell.Comment.Delete
    On Error GoTo 0
    If Len(Trim$(commentText)) = 0 Then Exit Sub

    targetCell.AddComment
    targetCell.Comment.Text commentText
    targetCell.Comment.Shape.TextFrame.Characters.Font.Name = "Bahnschrift"
    targetCell.Comment.Shape.TextFrame.AutoSize = True
End Sub

Public Sub ClearRangeComments(ByVal targetRange As Range)
    Dim someCell As Range

    If targetRange Is Nothing Then Exit Sub
    For Each someCell In targetRange.Cells
        On Error Resume Next
        If Not someCell.Comment Is Nothing Then someCell.Comment.Delete
        On Error GoTo 0
    Next someCell
End Sub

Public Sub UpdateDashboardStatus(ByVal messageText As String, Optional ByVal fillColor As Long = 0)
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    SetConfigValue NAME_STATUS, messageText
    With ws.Range("B11:H13")
        .MergeCells = True
        .Value = messageText
        .Interior.Color = IIf(fillColor = 0, RGB(255, 255, 255), fillColor)
        .Font.Color = COLOR_TEXT_DARK
        .Font.Name = "Bahnschrift SemiBold"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.Color = COLOR_BORDER
    End With
    Application.StatusBar = messageText
End Sub

Public Sub UpdateLastRefreshStamp()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    With ws.Range("I11:M13")
        .MergeCells = True
        .Value = "Last refresh: " & Format$(Now, "m/d/yyyy h:mm AM/PM")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Bahnschrift"
        .Interior.Color = COLOR_PANEL_BG
        .Borders.Color = COLOR_BORDER
    End With
End Sub

Public Sub RefreshDashboardSelectionCaptions()
    Dim ws As Worksheet
    Dim lines As String

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    lines = "Estimate Date: " & DescribeDateRange(GetConfigDate(NAME_ESTIMATE_DATE_FROM), GetConfigDate(NAME_ESTIMATE_DATE_TO)) & vbCrLf
    lines = lines & "Delivery Date: " & DescribeDateRange(GetConfigDate(NAME_DELIVERY_DATE_FROM), GetConfigDate(NAME_DELIVERY_DATE_TO)) & vbCrLf
    lines = lines & "Actual Work Date: " & DescribeDateRange(GetConfigDate(NAME_ACTUAL_WORK_DATE_FROM), GetConfigDate(NAME_ACTUAL_WORK_DATE_TO)) & vbCrLf
    lines = lines & "Customer Like: " & DescribeTextFilter(GetConfigText(NAME_CUSTOMER_LIKE)) & vbCrLf
    lines = lines & "Work Center Like: " & DescribeTextFilter(GetConfigText(NAME_WORK_CENTER_LIKE)) & vbCrLf
    lines = lines & "Employee Like: " & DescribeTextFilter(GetConfigText(NAME_EMPLOYEE_LIKE))

    With ws.Range("B26:M29")
        .MergeCells = True
        .Value = lines
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Interior.Color = COLOR_PANEL_BG
        .Borders.Color = COLOR_BORDER
        .Font.Name = "Bahnschrift"
    End With
End Sub

Public Sub SetupDashboardVisuals()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    ws.Cells.Clear
    DeleteShapesWithPrefix ws, "ppr_"

    With ws
        .Cells.Font.Name = "Bahnschrift"
        .Cells.Font.Color = COLOR_TEXT_DARK
        .Cells.Interior.Color = RGB(255, 252, 246)
        .Columns("A").ColumnWidth = 2.5
        .Columns("B:N").ColumnWidth = 12.5
        .Rows("1:60").RowHeight = 22

        .Range("B2:N3").Merge
        .Range("B2").Value = REPORT_TITLE
        .Range("B2").Interior.Color = COLOR_TITLE_BG
        .Range("B2").Font.Color = RGB(255, 255, 255)
        .Range("B2").Font.Size = 22
        .Range("B2").Font.Bold = True
        .Range("B2").HorizontalAlignment = xlLeft
        .Range("B2").VerticalAlignment = xlCenter

        .Range("B4:N4").Merge
        .Range("B4").Value = "Control Deck"
        .Range("B4").Font.Bold = True
        .Range("B4").Font.Size = 11

        .Range("B10:N10").Merge
        .Range("B10").Value = "Run lane: shape controls on top, scoreboard in the middle, raw detail on demand."
        .Range("B10").HorizontalAlignment = xlLeft
        .Range("B10").Interior.Color = COLOR_PANEL_BG
        .Range("B10").Borders.Color = COLOR_BORDER

        .Range("B15:N15").Merge
        .Range("B15").Value = "KPI Arena"
        .Range("B15").Font.Bold = True

        .Range("B25:N25").Merge
        .Range("B25").Value = "Active Filters"
        .Range("B25").Font.Bold = True

        .Range("B30:G30").Merge
        .Range("B30").Value = "Work Center Scoreboard"
        .Range("B30").Font.Bold = True
        .Range("B30").Interior.Color = COLOR_PANEL_BG
        .Range("B30").Borders.Color = COLOR_BORDER

        .Range("I30:N30").Merge
        .Range("I30").Value = "Employee Scoreboard"
        .Range("I30").Font.Bold = True
        .Range("I30").Interior.Color = COLOR_PANEL_BG
        .Range("I30").Borders.Color = COLOR_BORDER
    End With

    AddButtonOnRange ws, SHAPE_ESTIMATE_DATES, "Estimate Dates", "EditEstimateDateRange", ws.Range("B5:C6"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_DELIVERY_DATES, "Delivery Dates", "EditDeliveryDateRange", ws.Range("D5:E6"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_ACTUAL_DATES, "Actual Dates", "EditActualWorkDateRange", ws.Range("F5:G6"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_CUSTOMER_FILTER, "Customer", "EditCustomerLike", ws.Range("H5:I6"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_WORKCENTER_FILTER, "Work Center", "EditWorkCenterLike", ws.Range("J5:K6"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_EMPLOYEE_FILTER, "Employee", "EditEmployeeLike", ws.Range("L5:M6"), COLOR_FILTER_BG

    AddButtonOnRange ws, SHAPE_REFRESH, "Refresh", "RefreshReport", ws.Range("B8:C9"), COLOR_ACTION_BG
    AddButtonOnRange ws, SHAPE_CLEAR, "Clear Filters", "ClearAllFilters", ws.Range("D8:E9"), COLOR_ACCENT_BG
    AddButtonOnRange ws, SHAPE_CONFIG, "Config", "GoToConfigurationSheet", ws.Range("F8:G9"), COLOR_PANEL_BG
    AddButtonOnRange ws, SHAPE_DASHBOARD, "Dashboard", "GoToDashboardSheet", ws.Range("H8:I9"), COLOR_PANEL_BG
    AddButtonOnRange ws, SHAPE_DETAIL, "Detail Grid", "GoToDetailSheet", ws.Range("J8:K9"), COLOR_PANEL_BG
    AddButtonOnRange ws, SHAPE_QUERY_LOG, "Query Log", "GoToQueryLogSheet", ws.Range("L8:M9"), COLOR_PANEL_BG

    SetupKpiCard ws.Range("B17:D20"), "Query Rows", "-"
    SetupKpiCard ws.Range("E17:G20"), "Distinct Jobs", "-"
    SetupKpiCard ws.Range("H17:J20"), "Customers", "-"
    SetupKpiCard ws.Range("K17:M20"), "Estimated Hours", "-"
    SetupKpiCard ws.Range("B21:D24"), "Actual Hours", "-"
    SetupKpiCard ws.Range("E21:G24"), "Variance Hours", "-"
    SetupKpiCard ws.Range("H21:J24"), "Over Target Jobs", "-"
    SetupKpiCard ws.Range("K21:M24"), "Avg Ratio", "-"

    UpdateDashboardMetrics Empty
    RefreshDashboardSelectionCaptions
    ApplyDashboardComments
End Sub

Public Sub UpdateDashboardMetrics(ByVal dataMatrix As Variant)
    Dim ws As Worksheet
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim jobIds As Object
    Dim customerIds As Object
    Dim jobStats As Object
    Dim jobKey As Variant
    Dim jobDetail As Object
    Dim totalEstimated As Double
    Dim totalActual As Double
    Dim totalVariance As Double
    Dim ratioCount As Long
    Dim avgRatio As Double
    Dim overTargetJobs As Long

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    Set jobIds = CreateObject("Scripting.Dictionary")
    Set customerIds = CreateObject("Scripting.Dictionary")
    Set jobStats = CreateObject("Scripting.Dictionary")
    rowCount = MatrixRowCount(dataMatrix)

    For rowIndex = 1 To rowCount
        jobKey = NzText(dataMatrix(rowIndex, 1))
        If Len(jobKey) > 0 Then jobIds(jobKey) = True
        If Len(NzText(dataMatrix(rowIndex, 4))) > 0 Then customerIds(NzText(dataMatrix(rowIndex, 4))) = True

        totalEstimated = totalEstimated + NzNumber(dataMatrix(rowIndex, 10)) / 60#
        totalActual = totalActual + NzNumber(dataMatrix(rowIndex, 11)) / 60#
        totalVariance = totalVariance + NzNumber(dataMatrix(rowIndex, 14)) / 60#
        If NzNumber(dataMatrix(rowIndex, 15)) <> 0 Then
            avgRatio = avgRatio + NzNumber(dataMatrix(rowIndex, 15))
            ratioCount = ratioCount + 1
        End If

        If Len(jobKey) > 0 Then
            If Not jobStats.Exists(jobKey) Then
                Set jobDetail = CreateObject("Scripting.Dictionary")
                jobDetail("Estimated") = 0#
                jobDetail("Actual") = 0#
                Set jobStats(jobKey) = jobDetail
            End If
            Set jobDetail = jobStats(jobKey)
            jobDetail("Estimated") = NzNumber(jobDetail("Estimated")) + NzNumber(dataMatrix(rowIndex, 10))
            jobDetail("Actual") = NzNumber(jobDetail("Actual")) + NzNumber(dataMatrix(rowIndex, 11))
        End If
    Next rowIndex

    For Each jobKey In jobStats.Keys
        Set jobDetail = jobStats(jobKey)
        If NzNumber(jobDetail("Actual")) > NzNumber(jobDetail("Estimated")) Then overTargetJobs = overTargetJobs + 1
    Next jobKey

    If ratioCount > 0 Then avgRatio = avgRatio / ratioCount

    UpdateKpiCard ws.Range("B17:D20"), "Query Rows", Format$(rowCount, "#,##0")
    UpdateKpiCard ws.Range("E17:G20"), "Distinct Jobs", Format$(jobIds.Count, "#,##0")
    UpdateKpiCard ws.Range("H17:J20"), "Customers", Format$(customerIds.Count, "#,##0")
    UpdateKpiCard ws.Range("K17:M20"), "Estimated Hours", Format$(totalEstimated, "#,##0.00")
    UpdateKpiCard ws.Range("B21:D24"), "Actual Hours", Format$(totalActual, "#,##0.00")
    UpdateKpiCard ws.Range("E21:G24"), "Variance Hours", Format$(totalVariance, "#,##0.00")
    UpdateKpiCard ws.Range("H21:J24"), "Over Target Jobs", Format$(overTargetJobs, "#,##0")
    UpdateKpiCard ws.Range("K21:M24"), "Avg Ratio", Format$(avgRatio, "0.00x")
End Sub

Public Function EnsureListObject(ByVal ws As Worksheet, ByVal tableName As String, ByVal anchorAddress As String, ByVal headers As Variant) As ListObject
    Dim lo As ListObject
    Dim totalHeaders As Long
    Dim index As Long
    Dim headerRange As Range

    totalHeaders = HeaderCount(headers)
    Set headerRange = ws.Range(anchorAddress).Resize(1, totalHeaders)

    For index = 1 To totalHeaders
        headerRange.Cells(1, index).Value = headers(index)
    Next index

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, headerRange, , xlYes)
        lo.Name = tableName
        lo.TableStyle = "TableStyleMedium2"
    Else
        lo.Resize ws.Range(anchorAddress).Resize(Application.Max(2, lo.Range.Rows.Count), totalHeaders)
    End If

    Set EnsureListObject = lo
End Function

Public Sub WriteMatrixToTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal anchorAddress As String, ByVal headers As Variant, ByVal matrix As Variant)
    Dim lo As ListObject
    Dim rowCount As Long
    Dim colCount As Long
    Dim targetRange As Range
    Dim existingBody As Range

    Set lo = EnsureListObject(ws, tableName, anchorAddress, headers)
    rowCount = MatrixRowCount(matrix)
    colCount = HeaderCount(headers)
    Set existingBody = lo.DataBodyRange

    If rowCount <= 0 Then
        If Not existingBody Is Nothing Then existingBody.Delete
        Exit Sub
    End If

    If Not existingBody Is Nothing Then existingBody.ClearContents
    lo.Resize ws.Range(anchorAddress).Resize(rowCount + 1, colCount)
    Set targetRange = ws.Range(anchorAddress).Offset(1, 0).Resize(rowCount, colCount)
    targetRange.Value = matrix
End Sub

Public Sub ApplyPackingTableFormatting()
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = EnsureWorksheet(SHEET_PACKING_DETAIL)
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_PACKING_DETAIL)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    lo.HeaderRowRange.Font.Color = RGB(255, 255, 255)
    lo.HeaderRowRange.Font.Bold = True
    lo.Range.EntireColumn.AutoFit

    On Error Resume Next
    lo.ListColumns("Estimate Date").DataBodyRange.NumberFormat = "m/d/yyyy"
    lo.ListColumns("Delivery Date").DataBodyRange.NumberFormat = "m/d/yyyy"
    lo.ListColumns("Actual Work Date").DataBodyRange.NumberFormat = "m/d/yyyy"
    lo.ListColumns("Order Quantity").DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("Estimated Packing Minutes").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Actual Packing Minutes").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Total Actual Packing Minutes").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Share of Job Actual Time").DataBodyRange.NumberFormat = "0.00%"
    lo.ListColumns("Packing Minutes Variance").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Packing Time Ratio").DataBodyRange.NumberFormat = "0.00x"
    On Error GoTo 0
End Sub

Public Sub ApplySummaryTableFormatting(ByVal ws As Worksheet, ByVal tableName As String)
    Dim lo As ListObject

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    lo.HeaderRowRange.Font.Color = RGB(255, 255, 255)
    lo.HeaderRowRange.Font.Bold = True
    lo.Range.EntireColumn.AutoFit
    On Error Resume Next
    lo.ListColumns("Rows").DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("Jobs").DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("Est Hrs").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Act Hrs").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Var Hrs").DataBodyRange.NumberFormat = "#,##0.00"
    On Error GoTo 0
End Sub

Public Sub ApplyDashboardComments()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    SetCellComment ws.Range("B17"), "Query Rows" & vbLf & "Total rows returned by the active packing query."
    SetCellComment ws.Range("E17"), "Distinct Jobs" & vbLf & "Unique Job ID count in the active result set."
    SetCellComment ws.Range("H17"), "Customers" & vbLf & "Unique customer count in the active result set."
    SetCellComment ws.Range("K17"), "Estimated Hours" & vbLf & "Sum of Estimated Packing Minutes divided by 60."
    SetCellComment ws.Range("B21"), "Actual Hours" & vbLf & "Sum of Actual Packing Minutes divided by 60."
    SetCellComment ws.Range("E21"), "Variance Hours" & vbLf & "Sum of Packing Minutes Variance divided by 60."
    SetCellComment ws.Range("H21"), "Over Target Jobs" & vbLf & "Distinct jobs where actual packing minutes exceed estimated packing minutes."
    SetCellComment ws.Range("K21"), "Avg Ratio" & vbLf & "Average Packing Time Ratio across rows with a non-zero ratio."
    SetCellComment ws.Range("B25"), "Active Filters" & vbLf & "Dashboard caption of all current filter values. Blank values behave like SQL NULL."
    SetCellComment ws.Range("B30"), "Work Center Scoreboard" & vbLf & "Aggregated by work center with rows, jobs, estimated hours, actual hours, and variance hours."
    SetCellComment ws.Range("I30"), "Employee Scoreboard" & vbLf & "Aggregated by employee with rows, jobs, estimated hours, actual hours, and variance hours."
End Sub

Private Sub SetupKpiCard(ByVal targetRange As Range, ByVal titleText As String, ByVal valueText As String)
    With targetRange
        .Merge
        .Interior.Color = COLOR_CARD_BG
        .Borders.Color = COLOR_BORDER
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    UpdateKpiCard targetRange, titleText, valueText
End Sub

Private Sub UpdateKpiCard(ByVal targetRange As Range, ByVal titleText As String, ByVal valueText As String)
    Dim displayText As String

    displayText = titleText & vbLf & valueText
    With targetRange
        .Value = displayText
        .Font.Name = "Bahnschrift"
        .Font.Size = 11
        .Font.Bold = False
        .Characters(1, Len(titleText)).Font.Bold = True
        .Characters(1, Len(titleText)).Font.Size = 10
        If Len(valueText) > 0 Then
            .Characters(Len(titleText) + 2, Len(valueText)).Font.Bold = True
            .Characters(Len(titleText) + 2, Len(valueText)).Font.Size = 18
        End If
    End With
End Sub

Private Function DescribeDateRange(ByVal fromValue As Variant, ByVal toValue As Variant) As String
    Dim fromText As String
    Dim toText As String

    If IsEmpty(fromValue) Then fromText = "Any" Else fromText = Format$(CDate(fromValue), "m/d/yyyy")
    If IsEmpty(toValue) Then toText = "Any" Else toText = Format$(CDate(toValue), "m/d/yyyy")
    DescribeDateRange = fromText & " to " & toText
End Function

Private Function DescribeTextFilter(ByVal valueText As String) As String
    If Len(Trim$(valueText)) = 0 Then
        DescribeTextFilter = "Any"
    Else
        DescribeTextFilter = valueText
    End If
End Function
