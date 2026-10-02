Attribute VB_Name = "modValueSelector"
Option Explicit

Private Const VALUE_PREFIX As String = "pprVal_"

Private gValuePopupSheet As String
Private gValueFilterText As String
Private gValueRangeName As String
Private gValueAnchorShapeName As String
Private gValueLabelText As String
Private gValuePluralText As String
Private gValueColumnExpression As String
Private gValueSourceColumnIndex As Long
Private gValueItems() As String
Private gValueSelected() As Boolean
Private gValueItemCount As Long

Public Sub SelectCustomerFilter()
    OpenValueSelector "Customer", "Customers", NAME_CUSTOMER_LIKE, SHAPE_CUSTOMER_FILTER, "v.CustomerName", 4
End Sub

Public Sub SelectWorkCenterFilter()
    OpenValueSelector "Work Center", "Work Centers", NAME_WORK_CENTER_LIKE, SHAPE_WORKCENTER_FILTER, "v.WorkCenterName", 8
End Sub

Public Sub SelectEmployeeFilter()
    OpenValueSelector "Employee", "Employees", NAME_EMPLOYEE_LIKE, SHAPE_EMPLOYEE_FILTER, "v.EmployeeName", 9
End Sub

Private Sub OpenValueSelector(ByVal labelText As String, ByVal pluralText As String, ByVal rangeName As String, ByVal anchorShapeName As String, ByVal columnExpression As String, ByVal sourceColumnIndex As Long)
    Dim ws As Worksheet

    On Error GoTo PickerFail
    Set ws = EnsureWorksheet(SHEET_DASHBOARD)

    gValueLabelText = labelText
    gValuePluralText = pluralText
    gValueRangeName = rangeName
    gValueAnchorShapeName = anchorShapeName
    gValueColumnExpression = columnExpression
    gValueSourceColumnIndex = sourceColumnIndex
    gValuePopupSheet = ws.Name
    gValueFilterText = vbNullString

    LoadValueListCache
    AppendInteractionLog "Select" & Replace(labelText, " ", vbNullString), "Opened", "Opening value picker", DescribeCurrentFilters()
    ShowValuePicker ws, ws.Shapes(anchorShapeName)
    Exit Sub

PickerFail:
    AppendInteractionLog "Select" & Replace(labelText, " ", vbNullString), "Failure", SanitizedErrorMessage(Err.Description), DescribeCurrentFilters()
    UpdateDashboardStatus labelText & " picker failed: " & SanitizedErrorMessage(Err.Description), RGB(255, 199, 206)
End Sub

Private Sub LoadValueListCache()
    Dim selectedValues As Object
    Dim headers As Variant
    Dim data As Variant
    Dim rowIndex As Long

    Set selectedValues = ParseSelectedValues(GetConfigText(gValueRangeName))
    data = LoadDistinctPickerValues(headers)
    gValueItemCount = MatrixRowCount(data)

    If gValueItemCount <= 0 Then
        Erase gValueItems
        Erase gValueSelected
        Exit Sub
    End If

    ReDim gValueItems(1 To gValueItemCount)
    ReDim gValueSelected(1 To gValueItemCount)
    For rowIndex = 1 To gValueItemCount
        gValueItems(rowIndex) = Trim$(NzText(data(rowIndex, 1)))
        gValueSelected(rowIndex) = selectedValues.Exists(LCase$(gValueItems(rowIndex)))
    Next rowIndex
End Sub

Private Function LoadDistinctPickerValues(ByRef headers As Variant) As Variant
    Dim cn As Object
    Dim cmd As Object

    On Error GoTo UseDetailFallback
    Set cn = OpenCermConnection()
    Set cmd = CreateTextCommand(cn, BuildDistinctPickerSql())
    LoadDistinctPickerValues = ExecuteCommandMatrix(cmd, headers)
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
    End If
    Set cmd = Nothing
    Set cn = Nothing
    Exit Function

UseDetailFallback:
    AppendInteractionLog "Load" & Replace(gValueLabelText, " ", vbNullString) & "PickerCache", "Fallback", SanitizedErrorMessage(Err.Description), DescribeCurrentFilters()
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State <> 0 Then cn.Close
    Set cmd = Nothing
    Set cn = Nothing
    On Error GoTo 0
    LoadDistinctPickerValues = LoadDistinctPickerValuesFromDetail(headers)
End Function

Private Function BuildDistinctPickerSql() As String
    Dim sqlLines As Collection

    Set sqlLines = New Collection
    sqlLines.Add "SET NOCOUNT ON;"
    sqlLines.Add "SELECT DISTINCT ValueText = LTRIM(RTRIM(COALESCE(CAST(" & gValueColumnExpression & " AS nvarchar(255)), N'')))"
    sqlLines.Add "FROM dbo.vw_stlPackingEstimatedVsActual AS v"
    sqlLines.Add "WHERE LEN(LTRIM(RTRIM(COALESCE(CAST(" & gValueColumnExpression & " AS nvarchar(255)), N'')))) > 0"
    sqlLines.Add "  AND (" & BuildDatePredicate("v.EstimateDate", GetConfigDate(NAME_ESTIMATE_DATE_FROM), GetConfigDate(NAME_ESTIMATE_DATE_TO)) & ")"
    sqlLines.Add "  AND (" & BuildDatePredicate("v.DeliveryDate", GetConfigDate(NAME_DELIVERY_DATE_FROM), GetConfigDate(NAME_DELIVERY_DATE_TO)) & ")"
    sqlLines.Add "  AND (" & BuildDatePredicate("v.ActualWorkDate", GetConfigDate(NAME_ACTUAL_WORK_DATE_FROM), GetConfigDate(NAME_ACTUAL_WORK_DATE_TO)) & ")"
    If gValueRangeName <> NAME_CUSTOMER_LIKE Then sqlLines.Add "  AND (" & BuildDiscreteFilterPredicate("v.CustomerName", GetConfigText(NAME_CUSTOMER_LIKE)) & ")"
    If gValueRangeName <> NAME_WORK_CENTER_LIKE Then sqlLines.Add "  AND (" & BuildDiscreteFilterPredicate("v.WorkCenterName", GetConfigText(NAME_WORK_CENTER_LIKE)) & ")"
    If gValueRangeName <> NAME_EMPLOYEE_LIKE Then sqlLines.Add "  AND (" & BuildDiscreteFilterPredicate("v.EmployeeName", GetConfigText(NAME_EMPLOYEE_LIKE)) & ")"
    sqlLines.Add "ORDER BY ValueText;"
    BuildDistinctPickerSql = JoinCollection(sqlLines, vbCrLf)
End Function

Private Function BuildDatePredicate(ByVal columnExpression As String, ByVal fromValue As Variant, ByVal toValue As Variant) As String
    Dim clauses As Collection

    Set clauses = New Collection
    If IsEmpty(fromValue) Then
        clauses.Add "1 = 1"
    Else
        clauses.Add columnExpression & " >= " & SqlDateLiteral(fromValue)
    End If

    If Not IsEmpty(toValue) Then
        clauses.Add columnExpression & " < DATEADD(DAY, 1, " & SqlDateLiteral(toValue) & ")"
    End If

    BuildDatePredicate = JoinCollection(clauses, " AND ")
End Function

Private Function BuildDiscreteFilterPredicate(ByVal columnExpression As String, ByVal filterText As String) As String
    Dim escapedValue As String
    Dim normalizedCsv As String

    If Len(Trim$(filterText)) = 0 Then
        BuildDiscreteFilterPredicate = "1 = 1"
        Exit Function
    End If

    escapedValue = Replace(filterText, "'", "''")
    If HasSqlWildcardSyntax(filterText) Then
        BuildDiscreteFilterPredicate = "COALESCE(CAST(" & columnExpression & " AS nvarchar(255)), N'') LIKE N'" & escapedValue & "'"
    Else
        normalizedCsv = Replace(escapedValue, ", ", ",")
        BuildDiscreteFilterPredicate = "CHARINDEX(N',' + LTRIM(RTRIM(COALESCE(CAST(" & columnExpression & " AS nvarchar(255)), N''))) + N',', N',' + N'" & normalizedCsv & "' + N',') > 0"
    End If
End Function

Private Function LoadDistinctPickerValuesFromDetail(ByRef headers As Variant) As Variant
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim valuesByKey As Object
    Dim rowRange As Range
    Dim itemValue As String
    Dim result() As Variant
    Dim rowIndex As Long
    Dim key As Variant

    headers = Array("ValueText")
    Set ws = EnsureWorksheet(SHEET_PACKING_DETAIL)
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_PACKING_DETAIL)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Set valuesByKey = CreateObject("Scripting.Dictionary")
    For Each rowRange In lo.DataBodyRange.Rows
        itemValue = Trim$(NzText(rowRange.Cells(1, gValueSourceColumnIndex).Value))
        If Len(itemValue) > 0 Then valuesByKey(LCase$(itemValue)) = itemValue
    Next rowRange
    If valuesByKey.Count = 0 Then Exit Function

    ReDim result(1 To valuesByKey.Count, 1 To 1)
    rowIndex = 0
    For Each key In valuesByKey.Keys
        rowIndex = rowIndex + 1
        result(rowIndex, 1) = valuesByKey(key)
    Next key
    SortSingleColumnMatrix result
    LoadDistinctPickerValuesFromDetail = result
End Function

Private Sub SortSingleColumnMatrix(ByRef matrix As Variant)
    Dim rowIndex As Long
    Dim compareIndex As Long
    Dim tempValue As String

    For rowIndex = 1 To MatrixRowCount(matrix) - 1
        For compareIndex = rowIndex + 1 To MatrixRowCount(matrix)
            If StrComp(CStr(matrix(rowIndex, 1)), CStr(matrix(compareIndex, 1)), vbTextCompare) > 0 Then
                tempValue = CStr(matrix(rowIndex, 1))
                matrix(rowIndex, 1) = matrix(compareIndex, 1)
                matrix(compareIndex, 1) = tempValue
            End If
        Next compareIndex
    Next rowIndex
End Sub

Private Sub ShowValuePicker(ByVal ws As Worksheet, ByVal anchorShape As Shape)
    Dim rowIndex As Long
    Dim top0 As Single
    Dim left0 As Single
    Dim panelWidth As Single
    Dim panelHeight As Single
    Dim itemHeight As Single
    Dim visibleCount As Long
    Dim yPos As Single
    Dim itemShape As Shape
    Dim bg As Shape

    DeleteShapesWithPrefix ws, VALUE_PREFIX
    left0 = anchorShape.Left
    top0 = anchorShape.Top + anchorShape.Height + 4
    panelWidth = 250
    itemHeight = 18

    For rowIndex = 1 To gValueItemCount
        If MatchFilterText(gValueItems(rowIndex), gValueFilterText) Then visibleCount = visibleCount + 1
    Next rowIndex
    If visibleCount = 0 Then visibleCount = 1

    panelHeight = 78 + (visibleCount * itemHeight)
    Set bg = ws.Shapes.AddShape(msoShapeRoundedRectangle, left0, top0, panelWidth, panelHeight)
    bg.Name = VALUE_PREFIX & "bg"
    SetPopupOuterShapeStyle bg

    Set itemShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left0 + 8, top0 + 8, panelWidth - 90, 16)
    With itemShape
        .Name = VALUE_PREFIX & "title"
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        SetShapeTextCenter itemShape, "Select " & gValuePluralText, 9, True
    End With

    AddPopupButton ws, VALUE_PREFIX & "filter", "Filter", ThisWorkbook.Name & "!ValuePickerFilter", left0 + panelWidth - 120, top0 + 6, 36, 16
    AddPopupButton ws, VALUE_PREFIX & "all", "All", ThisWorkbook.Name & "!ValuePickerSelectAll", left0 + panelWidth - 82, top0 + 6, 18, 16
    AddPopupButton ws, VALUE_PREFIX & "apply", "Apply", ThisWorkbook.Name & "!ValuePickerApply", left0 + panelWidth - 62, top0 + 6, 40, 16
    AddPopupButton ws, VALUE_PREFIX & "close", "x", ThisWorkbook.Name & "!ValuePickerClose", left0 + panelWidth - 20, top0 + 6, 14, 16

    yPos = top0 + 38
    Set itemShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left0 + 8, yPos, panelWidth - 16, 14)
    With itemShape
        .Name = VALUE_PREFIX & "hint"
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        SetShapeTextCenter itemShape, "Click to toggle, then Apply", 7, False
    End With
    yPos = yPos + 20

    If gValueItemCount = 0 Then
        Set itemShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, left0 + 8, yPos, panelWidth - 16, itemHeight)
        With itemShape
            .Name = VALUE_PREFIX & "empty"
            .Line.Visible = msoFalse
            .Fill.Visible = msoFalse
            SetShapeTextCenter itemShape, "No values available", 8, False
        End With
        Exit Sub
    End If

    For rowIndex = 1 To gValueItemCount
        If Len(gValueItems(rowIndex)) > 0 And MatchFilterText(gValueItems(rowIndex), gValueFilterText) Then
            Set itemShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, left0 + 6, yPos, panelWidth - 12, itemHeight)
            With itemShape
                .Name = VALUE_PREFIX & "item_" & CStr(rowIndex)
                .OnAction = ThisWorkbook.Name & "!ValuePickerSelect"
                .Fill.ForeColor.RGB = IIf(gValueSelected(rowIndex), RGB(255, 242, 204), RGB(255, 255, 255))
                .Line.ForeColor.RGB = RGB(220, 220, 220)
            End With
            SetShapeTextCenter itemShape, SelectionPrefix(gValueSelected(rowIndex)) & gValueItems(rowIndex), 8, False
            yPos = yPos + itemHeight
        End If
    Next rowIndex
End Sub

Private Sub AddPopupButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, ByVal macroName As String, ByVal left0 As Single, ByVal top0 As Single, ByVal width0 As Single, ByVal height0 As Single)
    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, left0, top0, width0, height0)
    shp.Name = shapeName
    shp.OnAction = macroName
    shp.Fill.ForeColor.RGB = RGB(242, 242, 242)
    shp.Line.ForeColor.RGB = COLOR_BORDER
    SetShapeTextCenter shp, caption, 8, True
End Sub

Public Sub ValuePickerFilter()
    Dim userInput As Variant
    Dim ws As Worksheet

    userInput = Application.InputBox("Type " & LCase$(gValueLabelText) & " filter text:", "Filter " & gValuePluralText, gValueFilterText, Type:=2)
    If userInput = False Then Exit Sub
    gValueFilterText = CStr(userInput)
    Set ws = EnsureWorksheet(PopupSheetOrDashboard())
    AppendInteractionLog "ValuePickerFilter", "Success", gValueLabelText & " | " & gValueFilterText, DescribeCurrentFilters()
    ShowValuePicker ws, ws.Shapes(gValueAnchorShapeName)
End Sub

Public Sub ValuePickerSelectAll()
    Dim rowIndex As Long
    Dim ws As Worksheet

    For rowIndex = 1 To gValueItemCount
        gValueSelected(rowIndex) = False
    Next rowIndex
    Set ws = EnsureWorksheet(PopupSheetOrDashboard())
    AppendInteractionLog "ValuePickerSelectAll", "Success", gValueLabelText & " | All", DescribeCurrentFilters()
    ShowValuePicker ws, ws.Shapes(gValueAnchorShapeName)
End Sub

Public Sub ValuePickerClose()
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(PopupSheetOrDashboard())
    AppendInteractionLog "ValuePickerClose", "Success", "Closing " & gValueLabelText & " picker", DescribeCurrentFilters()
    DeleteShapesWithPrefix ws, VALUE_PREFIX
End Sub

Public Sub ValuePickerSelect()
    Dim rowIndex As Long
    Dim shapeName As String
    Dim ws As Worksheet

    shapeName = CStr(Application.Caller)
    rowIndex = CLng(Mid$(shapeName, Len(VALUE_PREFIX & "item_") + 1))
    gValueSelected(rowIndex) = Not gValueSelected(rowIndex)
    Set ws = EnsureWorksheet(PopupSheetOrDashboard())
    ShowValuePicker ws, ws.Shapes(gValueAnchorShapeName)
End Sub

Public Sub ValuePickerApply()
    Dim rowIndex As Long
    Dim selectedValues As String
    Dim selectedCount As Long
    Dim ws As Worksheet

    For rowIndex = 1 To gValueItemCount
        If gValueSelected(rowIndex) Then
            AppendCsvValue selectedValues, gValueItems(rowIndex)
            selectedCount = selectedCount + 1
        End If
    Next rowIndex

    SetConfigValue gValueRangeName, selectedValues
    Set ws = EnsureWorksheet(PopupSheetOrDashboard())
    DeleteShapesWithPrefix ws, VALUE_PREFIX
    RefreshDashboardSelectionCaptions
    UpdateDashboardStatus gValueLabelText & " filter updated. Click Refresh to apply.", RGB(255, 255, 255)
    AppendInteractionLog "Apply" & Replace(gValueLabelText, " ", vbNullString) & "Filter", "Success", BuildSelectedCaption(selectedValues, selectedCount, gValuePluralText), DescribeCurrentFilters()
End Sub

Public Sub ValuePickerSelectFirstVisible()
    Dim rowIndex As Long

    For rowIndex = 1 To gValueItemCount
        If MatchFilterText(gValueItems(rowIndex), gValueFilterText) Then
            gValueSelected(rowIndex) = Not gValueSelected(rowIndex)
            Exit Sub
        End If
    Next rowIndex
End Sub

Public Function DescribeDiscreteFilter(ByVal valueText As String, ByVal pluralLabel As String) As String
    Dim values As Variant
    Dim count As Long

    If Len(Trim$(valueText)) = 0 Then
        DescribeDiscreteFilter = "Any"
    ElseIf HasSqlWildcardSyntax(valueText) Then
        DescribeDiscreteFilter = valueText
    ElseIf InStr(1, valueText, ",", vbBinaryCompare) > 0 Then
        values = Split(valueText, ",")
        count = UBound(values) - LBound(values) + 1
        DescribeDiscreteFilter = CStr(count) & " " & pluralLabel
    Else
        DescribeDiscreteFilter = valueText
    End If
End Function

Private Function LoadCsvKeySet(ByVal csvText As String) As Object
    Dim result As Object
    Dim values As Variant
    Dim index As Long
    Dim itemText As String

    Set result = CreateObject("Scripting.Dictionary")
    If Len(Trim$(csvText)) = 0 Then
        Set LoadCsvKeySet = result
        Exit Function
    End If

    values = Split(csvText, ",")
    For index = LBound(values) To UBound(values)
        itemText = Trim$(CStr(values(index)))
        If Len(itemText) > 0 Then result(LCase$(itemText)) = True
    Next index
    Set LoadCsvKeySet = result
End Function

Private Function ParseSelectedValues(ByVal storedValue As String) As Object
    If HasSqlWildcardSyntax(storedValue) Then
        Set ParseSelectedValues = CreateObject("Scripting.Dictionary")
    Else
        Set ParseSelectedValues = LoadCsvKeySet(storedValue)
    End If
End Function

Private Function PopupSheetOrDashboard() As String
    If Len(Trim$(gValuePopupSheet)) = 0 Then
        PopupSheetOrDashboard = SHEET_DASHBOARD
    Else
        PopupSheetOrDashboard = gValuePopupSheet
    End If
End Function

Private Function MatchFilterText(ByVal sourceText As String, ByVal filterText As String) As Boolean
    If Len(Trim$(filterText)) = 0 Then
        MatchFilterText = True
    Else
        MatchFilterText = InStr(1, LCase$(sourceText), LCase$(Trim$(filterText)), vbTextCompare) > 0
    End If
End Function

Private Function HasSqlWildcardSyntax(ByVal valueText As String) As Boolean
    HasSqlWildcardSyntax = InStr(1, valueText, "%", vbBinaryCompare) > 0 Or _
        InStr(1, valueText, "_", vbBinaryCompare) > 0 Or _
        InStr(1, valueText, "[", vbBinaryCompare) > 0
End Function

Private Sub AppendCsvValue(ByRef csvText As String, ByVal valueText As String)
    If Len(Trim$(valueText)) = 0 Then Exit Sub
    If Len(csvText) = 0 Then
        csvText = Trim$(valueText)
    Else
        csvText = csvText & "," & Trim$(valueText)
    End If
End Sub

Private Function BuildSelectedCaption(ByVal selectedValues As String, ByVal selectedCount As Long, ByVal pluralLabel As String) As String
    If selectedCount = 0 Then
        BuildSelectedCaption = "All"
    ElseIf selectedCount = 1 Then
        BuildSelectedCaption = selectedValues
    Else
        BuildSelectedCaption = CStr(selectedCount) & " " & pluralLabel
    End If
End Function

Private Function SelectionPrefix(ByVal isSelected As Boolean) As String
    If isSelected Then
        SelectionPrefix = "[x] "
    Else
        SelectionPrefix = "[ ] "
    End If
End Function

Private Function JoinCollection(ByVal values As Collection, ByVal delimiter As String) As String
    Dim index As Long

    For index = 1 To values.Count
        If index > 1 Then JoinCollection = JoinCollection & delimiter
        JoinCollection = JoinCollection & CStr(values(index))
    Next index
End Function