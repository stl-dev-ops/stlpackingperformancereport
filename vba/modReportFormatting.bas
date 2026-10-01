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

Public Sub UpdateDashboardStatus(ByVal messageText As String, Optional ByVal fillColor As Long = 0)
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    SetConfigValue NAME_STATUS, messageText
    With ws.Range("B5:H6")
        .MergeCells = True
        .Value = messageText
        .Interior.Color = IIf(fillColor = 0, RGB(255, 255, 255), fillColor)
        .Font.Color = COLOR_TEXT_DARK
        .Font.Name = "Bahnschrift SemiBold"
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    Application.StatusBar = messageText
End Sub

Public Sub UpdateLastRefreshStamp()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    ws.Range("J5:M6").MergeCells = True
    With ws.Range("J5")
        .Value = "Last refresh: " & Format$(Now, "m/d/yyyy h:mm AM/PM")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "Bahnschrift"
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

    With ws.Range("B13:H18")
        .MergeCells = True
        .Value = lines
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Interior.Color = RGB(248, 248, 248)
        .Borders.Color = COLOR_BORDER
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
        .Columns("A").ColumnWidth = 3
        .Columns("B:M").ColumnWidth = 14
        .Rows("1:30").RowHeight = 22

        .Range("B2:M3").Merge
        .Range("B2").Value = REPORT_TITLE
        .Range("B2").Interior.Color = COLOR_TITLE_BG
        .Range("B2").Font.Color = RGB(255, 255, 255)
        .Range("B2").Font.Size = 20
        .Range("B2").Font.Bold = True
        .Range("B2").HorizontalAlignment = xlLeft
        .Range("B2").VerticalAlignment = xlCenter

        .Range("B8:F8").Value = Array("Rows", "Distinct Jobs", "Estimated Minutes", "Actual Minutes", "Variance Minutes")
        .Range("B8:F8").Font.Bold = True
        .Range("B9:F9").Interior.Color = RGB(248, 248, 248)
        .Range("B9:F9").Borders.Color = COLOR_BORDER

        .Range("J8:M8").Merge
        .Range("J8").Value = "How to use"
        .Range("J8").Font.Bold = True
        .Range("J9:M14").Merge
        .Range("J9").Value = "Set optional filters on the Configuration sheet, then click Refresh. Blank filter cells behave like NULL and do not restrict the query."
        .Range("J9").WrapText = True
        .Range("J9").VerticalAlignment = xlTop
        .Range("J9").Interior.Color = RGB(248, 248, 248)
        .Range("J9").Borders.Color = COLOR_BORDER

        .Range("B12:H12").Value = "Active filters"
        .Range("B12:H12").Font.Bold = True
    End With

    AddButtonOnRange ws, SHAPE_REFRESH, "Refresh", "RefreshReport", ws.Range("J2:L3"), COLOR_ACTION_BG
    AddButtonOnRange ws, SHAPE_CONFIG, "Configuration", "GoToConfigurationSheet", ws.Range("J5:L6"), COLOR_ACCENT_BG

    UpdateDashboardMetrics Empty
    RefreshDashboardSelectionCaptions
End Sub

Public Sub UpdateDashboardMetrics(ByVal dataMatrix As Variant)
    Dim ws As Worksheet
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim jobIds As Object
    Dim totalEstimated As Double
    Dim totalActual As Double
    Dim totalVariance As Double

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    Set jobIds = CreateObject("Scripting.Dictionary")
    rowCount = MatrixRowCount(dataMatrix)

    For rowIndex = 1 To rowCount
        jobIds(CStr(dataMatrix(rowIndex, 1))) = True
        totalEstimated = totalEstimated + NzNumber(dataMatrix(rowIndex, 10))
        totalActual = totalActual + NzNumber(dataMatrix(rowIndex, 11))
        totalVariance = totalVariance + NzNumber(dataMatrix(rowIndex, 14))
    Next rowIndex

    ws.Range("B9").Value = rowCount
    ws.Range("C9").Value = jobIds.Count
    ws.Range("D9").Value = totalEstimated
    ws.Range("E9").Value = totalActual
    ws.Range("F9").Value = totalVariance
    ws.Range("D9:F9").NumberFormat = "#,##0.00"
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

Private Function NzNumber(ByVal someValue As Variant, Optional ByVal defaultValue As Double = 0#) As Double
    If IsError(someValue) Or IsNull(someValue) Or IsEmpty(someValue) Or Not IsNumeric(someValue) Then
        NzNumber = defaultValue
    Else
        NzNumber = CDbl(someValue)
    End If
End Function
