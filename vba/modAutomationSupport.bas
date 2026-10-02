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

Public Function TestTableRowCount(ByVal sheetName As String, ByVal tableName As String) As Long
    Dim lo As ListObject

    On Error Resume Next
    Set lo = EnsureWorksheet(sheetName).ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
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

Public Function TestCommentExists(ByVal sheetName As String, ByVal cellAddress As String) As Boolean
    Dim ws As Worksheet

    Set ws = EnsureWorksheet(sheetName)
    On Error Resume Next
    TestCommentExists = Not ws.Range(cellAddress).Comment Is Nothing
    On Error GoTo 0
End Function