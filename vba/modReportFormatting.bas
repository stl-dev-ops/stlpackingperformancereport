Attribute VB_Name = "modReportFormatting"
Option Explicit

Private Const POPUP_CORNER_RADIUS_PX As Double = 6#

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

Public Sub SetShapeTextCenter(ByVal shp As Shape, ByVal caption As String, ByVal fontSize As Long, ByVal bold As Boolean)
    On Error Resume Next
    shp.TextFrame2.TextRange.Text = caption
    shp.TextFrame2.TextRange.Font.Size = fontSize
    shp.TextFrame2.TextRange.Font.Bold = IIf(bold, msoTrue, msoFalse)
    shp.TextFrame2.TextRange.Font.Name = "Bahnschrift"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COLOR_TEXT_DARK
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Color = COLOR_TEXT_DARK
    shp.TextFrame.Characters.Font.Name = "Bahnschrift"
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    On Error GoTo 0
End Sub

Public Sub SetPopupOuterShapeStyle(ByVal shp As Shape)
    Dim minDim As Double
    Dim adjustment As Double

    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.ForeColor.RGB = COLOR_BORDER
    shp.Line.Weight = 1
    shp.Placement = xlMoveAndSize

    minDim = shp.Width
    If shp.Height < minDim Then minDim = shp.Height
    If minDim <= 0 Then minDim = 1

    adjustment = POPUP_CORNER_RADIUS_PX / minDim
    If adjustment < 0# Then adjustment = 0#
    If adjustment > 0.5 Then adjustment = 0.5

    On Error Resume Next
    shp.Adjustments.Item(1) = adjustment
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
    SetConfigValue NAME_STATUS, messageText
    Application.StatusBar = messageText
End Sub

Public Sub UpdateLastRefreshStamp()
End Sub

Public Sub RefreshDashboardSelectionCaptions()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    If ShapeExists(ws, SHAPE_ESTIMATE_DATE_FROM) Then SetShapeButtonText ws.Shapes(SHAPE_ESTIMATE_DATE_FROM), "Est From" & vbLf & DescribeSingleDate(GetConfigDate(NAME_ESTIMATE_DATE_FROM))
    If ShapeExists(ws, SHAPE_ESTIMATE_DATE_TO) Then SetShapeButtonText ws.Shapes(SHAPE_ESTIMATE_DATE_TO), "Est To" & vbLf & DescribeSingleDate(GetConfigDate(NAME_ESTIMATE_DATE_TO))
    If ShapeExists(ws, SHAPE_DELIVERY_DATE_FROM) Then SetShapeButtonText ws.Shapes(SHAPE_DELIVERY_DATE_FROM), "Del From" & vbLf & DescribeSingleDate(GetConfigDate(NAME_DELIVERY_DATE_FROM))
    If ShapeExists(ws, SHAPE_DELIVERY_DATE_TO) Then SetShapeButtonText ws.Shapes(SHAPE_DELIVERY_DATE_TO), "Del To" & vbLf & DescribeSingleDate(GetConfigDate(NAME_DELIVERY_DATE_TO))
    If ShapeExists(ws, SHAPE_ACTUAL_DATE_FROM) Then SetShapeButtonText ws.Shapes(SHAPE_ACTUAL_DATE_FROM), "Act From" & vbLf & DescribeSingleDate(GetConfigDate(NAME_ACTUAL_WORK_DATE_FROM))
    If ShapeExists(ws, SHAPE_ACTUAL_DATE_TO) Then SetShapeButtonText ws.Shapes(SHAPE_ACTUAL_DATE_TO), "Act To" & vbLf & DescribeSingleDate(GetConfigDate(NAME_ACTUAL_WORK_DATE_TO))
    If ShapeExists(ws, SHAPE_CUSTOMER_FILTER) Then SetShapeButtonText ws.Shapes(SHAPE_CUSTOMER_FILTER), "Customer" & vbLf & DescribeTextFilter(GetConfigText(NAME_CUSTOMER_LIKE), "Customers")
    If ShapeExists(ws, SHAPE_WORKCENTER_FILTER) Then SetShapeButtonText ws.Shapes(SHAPE_WORKCENTER_FILTER), "Work Center" & vbLf & DescribeTextFilter(GetConfigText(NAME_WORK_CENTER_LIKE), "Work Centers")
    If ShapeExists(ws, SHAPE_EMPLOYEE_FILTER) Then SetShapeButtonText ws.Shapes(SHAPE_EMPLOYEE_FILTER), "Employee" & vbLf & DescribeTextFilter(GetConfigText(NAME_EMPLOYEE_LIKE), "Employees")
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
        .Columns("B:O").ColumnWidth = 11.3
        .Rows("1:90").RowHeight = 22

        .Range("B2:O3").Merge
        .Range("B2").Value = REPORT_TITLE
        .Range("B2").Interior.Color = COLOR_TITLE_BG
        .Range("B2").Font.Color = RGB(255, 255, 255)
        .Range("B2").Font.Size = 22
        .Range("B2").Font.Bold = True
        .Range("B2").HorizontalAlignment = xlLeft
        .Range("B2").VerticalAlignment = xlCenter

        .Range("B4:O4").Merge
        .Range("B4").Value = "Control Deck"
        .Range("B4").Font.Bold = True
        .Range("B4").Font.Size = 11

        .Range("B10:N10").Merge
        .Range("B10").Value = "KPI Arena"
        .Range("B10").Font.Bold = True

        .Range("B15:O15").Merge
        .Range("B15").Value = "Packing Detail Grid"
        .Range("B15").Font.Bold = True
        .Range("B15").Interior.Color = COLOR_PANEL_BG
        .Range("B15").Borders.Color = COLOR_BORDER

        .Range("B30:H30").Merge
        .Range("B30").Value = "Work Center Scoreboard"
        .Range("B30").Font.Bold = True
        .Range("B30").Interior.Color = COLOR_PANEL_BG
        .Range("B30").Borders.Color = COLOR_BORDER

        .Range("J30:P30").Merge
        .Range("J30").Value = "Employee Scoreboard"
        .Range("J30").Font.Bold = True
        .Range("J30").Interior.Color = COLOR_PANEL_BG
        .Range("J30").Borders.Color = COLOR_BORDER
    End With

    AddButtonOnRange ws, SHAPE_ESTIMATE_DATE_FROM, "Est From", "SelectEstimateDateFrom", ws.Range("B5:C6"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_ESTIMATE_DATE_TO, "Est To", "SelectEstimateDateTo", ws.Range("D5:E6"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_DELIVERY_DATE_FROM, "Del From", "SelectDeliveryDateFrom", ws.Range("F5:G6"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_DELIVERY_DATE_TO, "Del To", "SelectDeliveryDateTo", ws.Range("H5:I6"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_ACTUAL_DATE_FROM, "Act From", "SelectActualWorkDateFrom", ws.Range("J5:K6"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_ACTUAL_DATE_TO, "Act To", "SelectActualWorkDateTo", ws.Range("L5:M6"), COLOR_FILTER_BG

    AddButtonOnRange ws, SHAPE_CUSTOMER_FILTER, "Customer", "SelectCustomerFilter", ws.Range("B8:D9"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_WORKCENTER_FILTER, "Work Center", "SelectWorkCenterFilter", ws.Range("E8:G9"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_EMPLOYEE_FILTER, "Employee", "SelectEmployeeFilter", ws.Range("H8:J9"), COLOR_FILTER_BG
    AddButtonOnRange ws, SHAPE_REFRESH, "Refresh", "RefreshReport", ws.Range("K8:L9"), COLOR_ACTION_BG
    AddButtonOnRange ws, SHAPE_CLEAR, "Clear Filters", "ClearAllFilters", ws.Range("M8:N9"), COLOR_ACCENT_BG

    SetupKpiCard ws.Range("B11:C14"), "Distinct Jobs", "-"
    SetupKpiCard ws.Range("D11:E14"), "Customers", "-"
    SetupKpiCard ws.Range("F11:G14"), "Estimated Hours", "-"
    SetupKpiCard ws.Range("H11:I14"), "Actual Hours", "-"
    SetupKpiCard ws.Range("J11:K14"), "Variance Hours", "-"
    SetupKpiCard ws.Range("L11:M14"), "Over Target Jobs", "-"
    SetupKpiCard ws.Range("N11:O14"), "Variance %", "-"

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
    Dim percentageText As String
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

    totalVariance = totalActual - totalEstimated
    percentageText = "N/A"
    If totalEstimated > 0 Then percentageText = Format$(PackingPercentage(totalActual, totalEstimated), "+0.00%;-0.00%;0.00%")

    UpdateKpiCard ws.Range("B11:C14"), "Distinct Jobs", Format$(jobIds.Count, "#,##0")
    UpdateKpiCard ws.Range("D11:E14"), "Customers", Format$(customerIds.Count, "#,##0")
    UpdateKpiCard ws.Range("F11:G14"), "Estimated Hours", Format$(totalEstimated, "#,##0.00")
    UpdateKpiCard ws.Range("H11:I14"), "Actual Hours", Format$(totalActual, "#,##0.00")
    UpdateKpiCard ws.Range("J11:K14"), "Variance Hours", Format$(totalVariance, "#,##0.00")
    UpdateKpiCard ws.Range("L11:M14"), "Over Target Jobs", Format$(overTargetJobs, "#,##0")
    UpdateKpiCard ws.Range("N11:O14"), "Variance %", percentageText
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

    If Not existingBody Is Nothing Then
        If tableName = TABLE_WORKCENTER_SUMMARY Or tableName = TABLE_EMPLOYEE_SUMMARY Then ClearRangeComments existingBody
    End If
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
    lo.ListColumns("Variance %").DataBodyRange.NumberFormat = "+0.00%;-0.00%;0.00%"
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
    lo.HeaderRowRange.Interior.Pattern = xlSolid
    lo.HeaderRowRange.Interior.Color = COLOR_TITLE_BG
    lo.HeaderRowRange.WrapText = True
    lo.HeaderRowRange.VerticalAlignment = xlCenter
    lo.HeaderRowRange.RowHeight = 32
    lo.Range.EntireColumn.AutoFit
    On Error Resume Next
    lo.ListColumns("Jobs").DataBodyRange.NumberFormat = "#,##0"
    lo.ListColumns("Est Hrs").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Act Hrs").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Var Hrs").DataBodyRange.NumberFormat = "#,##0.00"
    lo.ListColumns("Variance %").DataBodyRange.NumberFormat = "+0.00%;-0.00%;0.00%"
    On Error GoTo 0
End Sub

Public Sub WriteDashboardDetailGrid(ByVal ws As Worksheet, ByVal headers As Variant, ByVal matrix As Variant)
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim rowCount As Long
    Dim targetCell As Range
    Dim gridRange As Range

    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_DASHBOARD_DETAIL)
    On Error GoTo 0
    If Not lo Is Nothing Then lo.Unlist

    ws.Range("B16:AA28").UnMerge
    ws.Range("B16:AA28").Clear
    rowCount = MatrixRowCount(matrix)
    Set gridRange = ws.Range("B16").Resize(rowCount + 1, HeaderCount(headers) * 2)
    For rowIndex = 0 To rowCount
        For columnIndex = 1 To HeaderCount(headers)
            Set targetCell = ws.Cells(16 + rowIndex, 2 + (columnIndex - 1) * 2).Resize(1, 2)
            targetCell.Merge
            targetCell.Borders.Color = COLOR_BORDER
            targetCell.VerticalAlignment = xlCenter
            targetCell.WrapText = True
            If rowIndex = 0 Then
                targetCell.Value = headers(columnIndex)
                targetCell.Font.Bold = True
                targetCell.Font.Color = RGB(255, 255, 255)
                targetCell.Interior.Color = COLOR_TITLE_BG
            Else
                targetCell.Value = matrix(rowIndex, columnIndex)
                targetCell.Interior.Color = IIf(rowIndex Mod 2 = 0, COLOR_PANEL_BG, RGB(255, 255, 255))
            End If
        Next columnIndex
    Next rowIndex
    ThisWorkbook.Names.Add Name:=TABLE_DASHBOARD_DETAIL, RefersTo:="=" & gridRange.Address(True, True, xlA1, True)
    ws.Columns("P:AA").ColumnWidth = 11.3
    ws.Rows("16:28").RowHeight = 32
    ApplyDashboardDetailFormatting ws
End Sub

Public Sub ApplyDashboardDetailFormatting(ByVal ws As Worksheet)
    ws.Range("F17:K28").NumberFormat = "#,##0.00"
    ws.Range("L17:M28").NumberFormat = "#,##0"
    ws.Range("N17:O28").NumberFormat = "+0.00%;-0.00%;0.00%"
    ws.Range("T17:U28").NumberFormat = "#,##0"
    ws.Range("V17:W28").NumberFormat = "m/d/yyyy"
End Sub

Public Sub ApplyDashboardComments()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    SetCellComment ws.Range("B11"), "Distinct Jobs" & vbLf & "Unique Job ID count in the active result set."
    SetCellComment ws.Range("D11"), "Customers" & vbLf & "Unique customer count in the active result set."
    SetCellComment ws.Range("F11"), "Estimated Hours" & vbLf & "Job estimates allocated by each row's share of the job's lifetime actual minutes, then summed across ALL filtered rows and divided by 60. Estimate-only jobs retain their full estimate."
    SetCellComment ws.Range("H11"), "Actual Hours" & vbLf & "Sum of Actual Packing Minutes divided by 60."
    SetCellComment ws.Range("J11"), "Variance Hours" & vbLf & "Actual Hours minus allocated Estimated Hours. Positive = over estimate; negative = under estimate."
    SetCellComment ws.Range("L11"), "Over Target Jobs" & vbLf & "Distinct jobs where actual packing minutes exceed estimated packing minutes."
    SetCellComment ws.Range("N11"), "Variance %" & vbLf & "(Total Actual Hours - total allocated Estimated Hours) / total allocated Estimated Hours across ALL filtered rows, not an average of row percentages or the 12-row preview. Negative = under; positive = over; 0% = on estimate. Zero actual is -100%. N/A = no positive estimate. Actual-only hours are included in total actual."
    SetCellComment ws.Range("B15"), "Packing Detail Grid" & vbLf & "Dashboard preview of the active result set directly below the KPI arena, ordered to mirror the KPI subjects first."
    SetCellComment ws.Range("B30"), "Work Center Scoreboard" & vbLf & "All filtered rows grouped by work center. Estimates allocated by share of job actual time; variance = actual minus allocated estimate; variance % = group variance / group estimate. Negative = under; positive = over. N/A = no positive estimate. Cell tooltips list contributing jobs."
    SetCellComment ws.Range("J30"), "Employee Scoreboard" & vbLf & "All filtered rows grouped by employee. Estimates allocated by share of job actual time, not independent employee budgets. Variance = actual minus allocated estimate; variance % = group variance / group estimate. Negative = under; positive = over. N/A = no positive estimate. Cell tooltips list contributing jobs."
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

Private Function DescribeSingleDate(ByVal someDate As Variant) As String
    If IsEmpty(someDate) Then
        DescribeSingleDate = "Any"
    Else
        DescribeSingleDate = Format$(CDate(someDate), "mmm d, yyyy")
    End If
End Function

Private Function DescribeTextFilter(ByVal valueText As String, ByVal pluralLabel As String) As String
    DescribeTextFilter = DescribeDiscreteFilter(valueText, pluralLabel)
End Function
