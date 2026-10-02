Attribute VB_Name = "modDateSelector"
Option Explicit

Private Const CAL_PREFIX As String = "pprCal_"

Private gCalTargetName As String
Private gCalAnchorShapeName As String
Private gCalPairedTargetName As String
Private gCalDisplayMonth As Date

Public Sub SelectEstimateDateFrom()
    AppendInteractionLog "SelectEstimateDateFrom", "Opened", "Opening estimate-from picker", DescribeCurrentFilters()
    ShowCalendarPicker NAME_ESTIMATE_DATE_FROM, NAME_ESTIMATE_DATE_TO, SHAPE_ESTIMATE_DATE_FROM
End Sub

Public Sub SelectEstimateDateTo()
    AppendInteractionLog "SelectEstimateDateTo", "Opened", "Opening estimate-to picker", DescribeCurrentFilters()
    ShowCalendarPicker NAME_ESTIMATE_DATE_TO, NAME_ESTIMATE_DATE_FROM, SHAPE_ESTIMATE_DATE_TO
End Sub

Public Sub SelectDeliveryDateFrom()
    AppendInteractionLog "SelectDeliveryDateFrom", "Opened", "Opening delivery-from picker", DescribeCurrentFilters()
    ShowCalendarPicker NAME_DELIVERY_DATE_FROM, NAME_DELIVERY_DATE_TO, SHAPE_DELIVERY_DATE_FROM
End Sub

Public Sub SelectDeliveryDateTo()
    AppendInteractionLog "SelectDeliveryDateTo", "Opened", "Opening delivery-to picker", DescribeCurrentFilters()
    ShowCalendarPicker NAME_DELIVERY_DATE_TO, NAME_DELIVERY_DATE_FROM, SHAPE_DELIVERY_DATE_TO
End Sub

Public Sub SelectActualWorkDateFrom()
    AppendInteractionLog "SelectActualWorkDateFrom", "Opened", "Opening actual-from picker", DescribeCurrentFilters()
    ShowCalendarPicker NAME_ACTUAL_WORK_DATE_FROM, NAME_ACTUAL_WORK_DATE_TO, SHAPE_ACTUAL_DATE_FROM
End Sub

Public Sub SelectActualWorkDateTo()
    AppendInteractionLog "SelectActualWorkDateTo", "Opened", "Opening actual-to picker", DescribeCurrentFilters()
    ShowCalendarPicker NAME_ACTUAL_WORK_DATE_TO, NAME_ACTUAL_WORK_DATE_FROM, SHAPE_ACTUAL_DATE_TO
End Sub

Private Sub ShowCalendarPicker(ByVal configName As String, ByVal pairedConfigName As String, ByVal shapeName As String)
    Dim ws As Worksheet
    Dim initialDate As Date

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    initialDate = GetConfigDate(configName, Date)

    gCalTargetName = configName
    gCalPairedTargetName = pairedConfigName
    gCalAnchorShapeName = shapeName
    gCalDisplayMonth = DateSerial(Year(initialDate), Month(initialDate), 1)
    RenderCalendarPicker ws, ws.Shapes(shapeName)
End Sub

Public Sub CalendarPickerClickDay()
    Dim picked As Date

    If Not TryParseCalDayShapeName(CStr(Application.Caller), picked) Then Exit Sub
    ApplySelectedDate gCalTargetName, picked, False
End Sub

Public Sub CalendarPickerPrevMonth()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    gCalDisplayMonth = DateAdd("m", -1, gCalDisplayMonth)
    AppendInteractionLog "CalendarPickerPrevMonth", "Success", Format$(gCalDisplayMonth, "yyyy-mm"), DescribeCurrentFilters()
    RenderCalendarPicker ws, ws.Shapes(gCalAnchorShapeName)
End Sub

Public Sub CalendarPickerNextMonth()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    gCalDisplayMonth = DateAdd("m", 1, gCalDisplayMonth)
    AppendInteractionLog "CalendarPickerNextMonth", "Success", Format$(gCalDisplayMonth, "yyyy-mm"), DescribeCurrentFilters()
    RenderCalendarPicker ws, ws.Shapes(gCalAnchorShapeName)
End Sub

Public Sub CalendarPickerClose()
    AppendInteractionLog "CalendarPickerClose", "Success", "Closing date picker", DescribeCurrentFilters()
    DeleteShapesWithPrefix EnsureWorksheet(SHEET_DASHBOARD), CAL_PREFIX
End Sub

Public Sub ApplySelectedDate(ByVal configName As String, ByVal picked As Date, Optional ByVal shouldRefresh As Boolean = False)
    Dim ws As Worksheet

    SetConfigValue configName, DateValue(picked)
    gCalTargetName = configName
    EnforceDateRange
    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    DeleteShapesWithPrefix ws, CAL_PREFIX
    RefreshDashboardSelectionCaptions
    UpdateDashboardStatus "Date updated. Click Refresh to apply.", RGB(255, 255, 255)
    AppendInteractionLog "ApplySelectedDate", "Success", configName & "=" & Format$(picked, "yyyy-mm-dd"), DescribeCurrentFilters()
    If shouldRefresh Then RefreshReport
End Sub

Private Sub RenderCalendarPicker(ByVal ws As Worksheet, ByVal anchorShape As Shape)
    Dim pad As Single
    Dim headerH As Single
    Dim dowH As Single
    Dim cellW As Single
    Dim cellH As Single
    Dim totalW As Single
    Dim totalH As Single
    Dim left0 As Single
    Dim top0 As Single
    Dim firstOfMonth As Date
    Dim daysInMonth As Long
    Dim offset As Long
    Dim selectedDate As Date
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim dayNum As Long
    Dim shp As Shape
    Dim dows As Variant
    Dim thisDate As Date

    DeleteShapesWithPrefix ws, CAL_PREFIX
    pad = 8
    headerH = 22
    dowH = 16
    cellW = 28
    cellH = 20
    totalW = pad * 2 + cellW * 7
    totalH = pad * 2 + headerH + dowH + cellH * 6
    left0 = anchorShape.Left
    top0 = anchorShape.Top + anchorShape.Height + 4

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left0, top0, totalW, totalH)
    shp.Name = CAL_PREFIX & "bg"
    SetPopupOuterShapeStyle shp

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left0 + pad + 22, top0 + pad, totalW - (pad * 2 + 44), headerH)
    shp.Name = CAL_PREFIX & "month"
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
    SetShapeTextCenter shp, Format$(gCalDisplayMonth, "mmmm yyyy"), 10, True

    AddCalendarNav ws, CAL_PREFIX & "prev", "<", ThisWorkbook.Name & "!CalendarPickerPrevMonth", left0 + pad, top0 + pad + 2, 18, headerH - 4
    AddCalendarNav ws, CAL_PREFIX & "next", ">", ThisWorkbook.Name & "!CalendarPickerNextMonth", left0 + totalW - pad - 18, top0 + pad + 2, 18, headerH - 4
    AddCalendarNav ws, CAL_PREFIX & "close", "x", ThisWorkbook.Name & "!CalendarPickerClose", left0 + totalW - pad - 18, top0 + totalH - pad - 18, 18, 18

    dows = Array("S", "M", "T", "W", "T", "F", "S")
    For colIndex = 0 To 6
        Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left0 + pad + cellW * colIndex, top0 + pad + headerH, cellW, dowH)
        shp.Name = CAL_PREFIX & "dow_" & CStr(colIndex)
        shp.Line.Visible = msoFalse
        shp.Fill.Visible = msoFalse
        SetShapeTextCenter shp, CStr(dows(colIndex)), 8, True
    Next colIndex

    selectedDate = GetConfigDate(gCalTargetName, Date)
    firstOfMonth = gCalDisplayMonth
    daysInMonth = Day(DateSerial(Year(firstOfMonth), Month(firstOfMonth) + 1, 0))
    offset = Weekday(firstOfMonth, vbSunday) - 1

    dayNum = 1
    For rowIndex = 0 To 5
        For colIndex = 0 To 6
            If (rowIndex * 7 + colIndex) >= offset And dayNum <= daysInMonth Then
                thisDate = DateSerial(Year(firstOfMonth), Month(firstOfMonth), dayNum)
                Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left0 + pad + cellW * colIndex, top0 + pad + headerH + dowH + cellH * rowIndex, cellW, cellH)
                shp.Name = CAL_PREFIX & "day_" & Format$(thisDate, "yyyymmdd")
                shp.OnAction = ThisWorkbook.Name & "!CalendarPickerClickDay"
                shp.Line.ForeColor.RGB = RGB(220, 220, 220)
                shp.Fill.ForeColor.RGB = IIf(thisDate = selectedDate, RGB(198, 223, 255), RGB(255, 255, 255))
                SetShapeTextCenter shp, CStr(dayNum), 8, False
                dayNum = dayNum + 1
            End If
        Next colIndex
    Next rowIndex
End Sub

Private Sub AddCalendarNav(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, ByVal macroName As String, ByVal left0 As Single, ByVal top0 As Single, ByVal width0 As Single, ByVal height0 As Single)
    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left0, top0, width0, height0)
    shp.Name = shapeName
    shp.OnAction = macroName
    shp.Fill.ForeColor.RGB = RGB(242, 242, 242)
    shp.Line.ForeColor.RGB = COLOR_BORDER
    SetShapeTextCenter shp, caption, 9, True
End Sub

Private Function TryParseCalDayShapeName(ByVal shapeName As String, ByRef outDate As Date) As Boolean
    Dim token As String

    token = Mid$(shapeName, Len(CAL_PREFIX & "day_") + 1)
    If Len(token) <> 8 Then Exit Function
    outDate = DateSerial(CInt(Left$(token, 4)), CInt(Mid$(token, 5, 2)), CInt(Right$(token, 2)))
    TryParseCalDayShapeName = True
End Function

Private Sub EnforceDateRange()
    Dim selectedDate As Date
    Dim pairedDate As Variant

    selectedDate = GetConfigDate(gCalTargetName, Date)
    pairedDate = GetConfigDate(gCalPairedTargetName)
    If IsEmpty(pairedDate) Then Exit Sub

    Select Case gCalTargetName
        Case NAME_ESTIMATE_DATE_FROM, NAME_DELIVERY_DATE_FROM, NAME_ACTUAL_WORK_DATE_FROM
            If CDate(pairedDate) < selectedDate Then SetConfigValue gCalPairedTargetName, selectedDate
        Case Else
            If CDate(pairedDate) > selectedDate Then SetConfigValue gCalPairedTargetName, selectedDate
    End Select
End Sub