Attribute VB_Name = "modAutomationSupport"
Option Explicit

Public Function AutomationInvoke(ByVal actionName As String, Optional ByVal arg1 As String = "", Optional ByVal arg2 As String = "") As String
    On Error GoTo EH

    Select Case LCase$(Trim$(actionName))
        Case "new-run-id"
            AutomationInvoke = NewTestRunId(arg1)
            Exit Function
        Case "initialize"
            InitializeReportWorkbook
        Case "open-estimate-from"
            SelectEstimateDateFrom
        Case "open-estimate-to"
            SelectEstimateDateTo
        Case "open-delivery-from"
            SelectDeliveryDateFrom
        Case "open-delivery-to"
            SelectDeliveryDateTo
        Case "open-actual-from"
            SelectActualWorkDateFrom
        Case "open-actual-to"
            SelectActualWorkDateTo
        Case "open-customer"
            SelectCustomerFilter
        Case "open-workcenter"
            SelectWorkCenterFilter
        Case "open-employee"
            SelectEmployeeFilter
        Case "calendar-next-month"
            CalendarPickerNextMonth
        Case "calendar-prev-month"
            CalendarPickerPrevMonth
        Case "calendar-close"
            CalendarPickerClose
        Case "value-picker-select-first"
            ValuePickerSelectFirstVisible
        Case "value-picker-set-filter"
            ValuePickerSetFilterText arg1
        Case "value-picker-select-all"
            ValuePickerSelectAll
        Case "value-picker-select-all-visible"
            ValuePickerSelectAllVisible
        Case "value-picker-apply"
            ValuePickerApply
        Case "value-picker-close"
            ValuePickerClose
        Case "refresh"
            RefreshReport
            If InStr(1, TestDashboardStatus(), "failed", vbTextCompare) > 0 Then
                AutomationInvoke = "ERR:REFRESH:" & TestDashboardStatus()
                Exit Function
            End If
        Case "clear-filters"
            ClearAllFilters
        Case "goto-configuration"
            GoToConfigurationSheet
        Case "goto-dashboard"
            GoToDashboardSheet
        Case "goto-detail"
            GoToDetailSheet
        Case "goto-query-log"
            GoToQueryLogSheet
        Case "set-estimate-range"
            TestSetDateRangeRaw NAME_ESTIMATE_DATE_FROM, NAME_ESTIMATE_DATE_TO, arg1, arg2
        Case "set-delivery-range"
            TestSetDateRangeRaw NAME_DELIVERY_DATE_FROM, NAME_DELIVERY_DATE_TO, arg1, arg2
        Case "set-actual-range"
            TestSetDateRangeRaw NAME_ACTUAL_WORK_DATE_FROM, NAME_ACTUAL_WORK_DATE_TO, arg1, arg2
        Case "set-customer"
            TestSetTextFilterRaw NAME_CUSTOMER_LIKE, arg1
        Case "set-workcenter"
            TestSetTextFilterRaw NAME_WORK_CENTER_LIKE, arg1
        Case "set-employee"
            TestSetTextFilterRaw NAME_EMPLOYEE_LIKE, arg1
        Case Else
            Err.Raise vbObjectError + 303, "AutomationInvoke", "Unknown automation action: " & actionName
    End Select

    AutomationInvoke = "OK"
    Exit Function

EH:
    On Error Resume Next
    UpdateDashboardStatus "Automation error: " & SanitizedErrorMessage(Err.Description), RGB(255, 199, 206)
    AppendInteractionLog "AutomationInvoke", "Failure", actionName & " | " & Err.Number & ":" & Err.Description, arg1 & " | " & arg2
    AutomationInvoke = "ERR:" & CStr(Err.Number) & ":" & SanitizedErrorMessage(Err.Description)
End Function

Public Sub TestSetDateRangeRaw(ByVal fromName As String, ByVal toName As String, ByVal startText As String, ByVal endText As String)
    If Len(Trim$(startText)) = 0 Then
        SetConfigValue fromName, Empty
    Else
        SetConfigValue fromName, CDate(startText)
    End If

    If Len(Trim$(endText)) = 0 Then
        SetConfigValue toName, Empty
    Else
        SetConfigValue toName, CDate(endText)
    End If

    RefreshDashboardSelectionCaptions
End Sub

Public Sub TestSetTextFilterRaw(ByVal rangeName As String, ByVal valueText As String)
    SetConfigValue rangeName, Trim$(valueText)
    RefreshDashboardSelectionCaptions
End Sub

Public Function TestShapeExists(ByVal sheetName As String, ByVal shapeName As String) As Boolean
    TestShapeExists = ShapeExists(EnsureWorksheet(sheetName), shapeName)
End Function

Public Function TestShapeText(ByVal sheetName As String, ByVal shapeName As String) As String
    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = EnsureWorksheet(sheetName)
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then Exit Function
    TestShapeText = shp.TextFrame.Characters.Text
End Function

Public Function TestTableRowCount(ByVal sheetName As String, ByVal tableName As String) As Long
    Dim lo As ListObject
    Dim gridRange As Range

    On Error Resume Next
    Set lo = EnsureWorksheet(sheetName).ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then
        If tableName = TABLE_DASHBOARD_DETAIL Then
            Set gridRange = ThisWorkbook.Names(tableName).RefersToRange
            TestTableRowCount = gridRange.Rows.Count - 1
        End If
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        TestTableRowCount = 0
    Else
        TestTableRowCount = lo.DataBodyRange.Rows.Count
    End If
End Function

Public Function TestDashboardStatus() As String
    TestDashboardStatus = GetConfigText(NAME_STATUS)
End Function

Public Function TestConfigText(ByVal rangeName As String) As String
    TestConfigText = GetConfigText(rangeName)
End Function

Public Function TestCurrentSheetName() As String
    TestCurrentSheetName = ActiveSheet.Name
End Function

Public Function TestActiveCellAddress() As String
    TestActiveCellAddress = ActiveCell.Address(False, False)
End Function

Public Function TestActiveWindowScrollRow() As Long
    On Error Resume Next
    TestActiveWindowScrollRow = ActiveWindow.ScrollRow
    On Error GoTo 0
End Function

Public Function TestActiveWindowScrollColumn() As Long
    On Error Resume Next
    TestActiveWindowScrollColumn = ActiveWindow.ScrollColumn
    On Error GoTo 0
End Function

Public Function TestCommentExists(ByVal sheetName As String, ByVal cellAddress As String) As Boolean
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(sheetName)
    On Error Resume Next
    TestCommentExists = Not ws.Range(cellAddress).Comment Is Nothing
    On Error GoTo 0
End Function

Public Function TestCellText(ByVal sheetName As String, ByVal cellAddress As String) As String
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(sheetName)
    TestCellText = CStr(ws.Range(cellAddress).Value)
End Function

Public Function TestKpiNumericValue(ByVal cellAddress As String) As Double
    Dim rawText As String
    Dim parts As Variant
    Dim valueText As String

    rawText = CStr(EnsureWorksheet(SHEET_DASHBOARD).Range(cellAddress).Value)
    parts = Split(rawText, vbLf)
    valueText = Trim$(CStr(parts(UBound(parts))))
    valueText = Replace(valueText, ",", "")
    valueText = Replace(valueText, "x", "")
    valueText = Replace(valueText, "%", "")
    If IsNumeric(valueText) Then TestKpiNumericValue = CDbl(valueText)
End Function

Public Function TestDashboardLayoutIssueCount() As Long
    TestDashboardLayoutIssueCount = CountDashboardLayoutIssues(EnsureWorksheet(SHEET_DASHBOARD))
End Function

Public Function TestDashboardGridAligned() As Boolean
    Dim ws As Worksheet
    Dim gridRange As Range
    Dim kpiRange As Range
    Dim fieldRange As Range
    Dim fieldIndex As Long
    Dim rowIndex As Long

    Set ws = EnsureWorksheet(SHEET_DASHBOARD)
    Set gridRange = ThisWorkbook.Names(TABLE_DASHBOARD_DETAIL).RefersToRange
    For fieldIndex = 0 To 6
        Set kpiRange = ws.Cells(11, 2 + fieldIndex * 2).MergeArea
        For rowIndex = 16 To 16 + gridRange.Rows.Count - 1
            Set fieldRange = ws.Cells(rowIndex, 2 + fieldIndex * 2).MergeArea
            If fieldRange.Columns.Count <> 2 Then Exit Function
            If Abs(fieldRange.Left - kpiRange.Left) > 0.1 Then Exit Function
            If Abs(fieldRange.Width - kpiRange.Width) > 0.1 Then Exit Function
        Next rowIndex
    Next fieldIndex
    TestDashboardGridAligned = True
End Function

Private Function CountDashboardLayoutIssues(ByVal ws As Worksheet) As Long
    Dim i As Long
    Dim j As Long
    Dim shapeA As Shape
    Dim shapeB As Shape
    Dim listObjectA As ListObject
    Dim listObjectB As ListObject

    For i = 1 To ws.Shapes.Count
        Set shapeA = ws.Shapes(i)
        If IsAuditedDashboardShape(shapeA) Then
            For j = i + 1 To ws.Shapes.Count
                Set shapeB = ws.Shapes(j)
                If IsAuditedDashboardShape(shapeB) Then
                    If ShapesOverlap(shapeA, shapeB) Then CountDashboardLayoutIssues = CountDashboardLayoutIssues + 1
                End If
            Next j

            For Each listObjectA In ws.ListObjects
                If ShapeOverlapsRange(shapeA, listObjectA.Range) Then CountDashboardLayoutIssues = CountDashboardLayoutIssues + 1
            Next listObjectA
        End If
    Next i

    For i = 1 To ws.ListObjects.Count
        Set listObjectA = ws.ListObjects(i)
        For j = i + 1 To ws.ListObjects.Count
            Set listObjectB = ws.ListObjects(j)
            If Not Application.Intersect(listObjectA.Range, listObjectB.Range) Is Nothing Then
                CountDashboardLayoutIssues = CountDashboardLayoutIssues + 1
            End If
        Next j
    Next i
End Function

Private Function IsAuditedDashboardShape(ByVal shp As Shape) As Boolean
    IsAuditedDashboardShape = (Left$(LCase$(shp.Name), 7) = "ppr_btn")
End Function

Private Function ShapesOverlap(ByVal shapeA As Shape, ByVal shapeB As Shape) As Boolean
    ShapesOverlap = Not (shapeA.Left + shapeA.Width <= shapeB.Left Or _
                         shapeB.Left + shapeB.Width <= shapeA.Left Or _
                         shapeA.Top + shapeA.Height <= shapeB.Top Or _
                         shapeB.Top + shapeB.Height <= shapeA.Top)
End Function

Private Function ShapeOverlapsRange(ByVal shp As Shape, ByVal targetRange As Range) As Boolean
    ShapeOverlapsRange = Not (shp.Left + shp.Width <= targetRange.Left Or _
                              targetRange.Left + targetRange.Width <= shp.Left Or _
                              shp.Top + shp.Height <= targetRange.Top Or _
                              targetRange.Top + targetRange.Height <= shp.Top)
End Function