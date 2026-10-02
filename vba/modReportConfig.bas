Attribute VB_Name = "modReportConfig"
Option Explicit

Public Const REPORT_VERSION As String = "0.2.0"
Public Const REPORT_TITLE As String = "Packing Estimated vs Actual Report"

Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_CONFIGURATION As String = "Configuration"
Public Const SHEET_PACKING_DETAIL As String = "Packing Detail"
Public Const SHEET_QUERY_LOG As String = "Query Log"
Public Const SHEET_INTERACTION_LOG As String = "Interaction Log"
Public Const SHEET_TEST_RESULTS As String = "Test Results"

Public Const TABLE_PACKING_DETAIL As String = "tblPackingDetail"
Public Const TABLE_QUERY_LOG As String = "tblQueryLog"
Public Const TABLE_INTERACTION_LOG As String = "tblInteractionLog"
Public Const TABLE_TEST_RESULTS As String = "tblTestResults"
Public Const TABLE_WORKCENTER_SUMMARY As String = "tblWorkCenterSummary"
Public Const TABLE_EMPLOYEE_SUMMARY As String = "tblEmployeeSummary"

Public Const NAME_ESTIMATE_DATE_FROM As String = "cfgEstimateDateFrom"
Public Const NAME_ESTIMATE_DATE_TO As String = "cfgEstimateDateTo"
Public Const NAME_DELIVERY_DATE_FROM As String = "cfgDeliveryDateFrom"
Public Const NAME_DELIVERY_DATE_TO As String = "cfgDeliveryDateTo"
Public Const NAME_ACTUAL_WORK_DATE_FROM As String = "cfgActualWorkDateFrom"
Public Const NAME_ACTUAL_WORK_DATE_TO As String = "cfgActualWorkDateTo"
Public Const NAME_CUSTOMER_LIKE As String = "cfgCustomerLike"
Public Const NAME_WORK_CENTER_LIKE As String = "cfgWorkCenterLike"
Public Const NAME_EMPLOYEE_LIKE As String = "cfgEmployeeLike"
Public Const NAME_SERVER_NAME As String = "cfgServerName"
Public Const NAME_DATABASE_NAME As String = "cfgDatabaseName"
Public Const NAME_COMMAND_TIMEOUT As String = "cfgCommandTimeout"
Public Const NAME_CONNECTION_TIMEOUT As String = "cfgConnectionTimeout"
Public Const NAME_STATUS As String = "cfgStatus"
Public Const NAME_VERSION As String = "cfgWorkbookVersion"

Public Const SHAPE_ESTIMATE_DATE_FROM As String = "ppr_btnEstimateDateFrom"
Public Const SHAPE_ESTIMATE_DATE_TO As String = "ppr_btnEstimateDateTo"
Public Const SHAPE_DELIVERY_DATE_FROM As String = "ppr_btnDeliveryDateFrom"
Public Const SHAPE_DELIVERY_DATE_TO As String = "ppr_btnDeliveryDateTo"
Public Const SHAPE_ACTUAL_DATE_FROM As String = "ppr_btnActualDateFrom"
Public Const SHAPE_ACTUAL_DATE_TO As String = "ppr_btnActualDateTo"
Public Const SHAPE_CUSTOMER_FILTER As String = "ppr_btnCustomerFilter"
Public Const SHAPE_WORKCENTER_FILTER As String = "ppr_btnWorkCenterFilter"
Public Const SHAPE_EMPLOYEE_FILTER As String = "ppr_btnEmployeeFilter"
Public Const SHAPE_REFRESH As String = "ppr_btnRefresh"
Public Const SHAPE_CLEAR As String = "ppr_btnClear"
Public Const SHAPE_CONFIG As String = "ppr_btnConfig"
Public Const SHAPE_DASHBOARD As String = "ppr_btnDashboard"
Public Const SHAPE_DETAIL As String = "ppr_btnDetail"
Public Const SHAPE_QUERY_LOG As String = "ppr_btnQueryLog"

Public Const COLOR_BORDER As Long = 11184810
Public Const COLOR_TEXT_DARK As Long = 2368548
Public Const COLOR_TITLE_BG As Long = 2624032
Public Const COLOR_ACTION_BG As Long = 10395294
Public Const COLOR_ACCENT_BG As Long = 15790320
Public Const COLOR_FILTER_BG As Long = 14671839
Public Const COLOR_PANEL_BG As Long = 15923442
Public Const COLOR_CARD_BG As Long = 16514555
Public Const COLOR_ALERT_BG As Long = 13421823
Public Const COLOR_SUCCESS_BG As Long = 13434828

Public Function WorkbookHasName(ByVal rangeName As String) As Boolean
    On Error Resume Next
    WorkbookHasName = Len(ThisWorkbook.Names(rangeName).Name) > 0
    On Error GoTo 0
End Function

Public Function GetNamedRange(ByVal rangeName As String) As Range
    If WorkbookHasName(rangeName) Then
        Set GetNamedRange = ThisWorkbook.Names(rangeName).RefersToRange
    End If
End Function

Public Function EnsureWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

Public Function GetConfigText(ByVal rangeName As String, Optional ByVal defaultValue As String = "") As String
    Dim rng As Range
    Dim rawValue As String

    Set rng = GetNamedRange(rangeName)
    If rng Is Nothing Then
        GetConfigText = defaultValue
    ElseIf IsError(rng.Value) Or IsNull(rng.Value) Then
        GetConfigText = defaultValue
    Else
        rawValue = CStr(rng.Value)
        If Len(Trim$(rawValue)) = 0 Then
            GetConfigText = defaultValue
        Else
            GetConfigText = rawValue
        End If
    End If
End Function

Public Function GetConfigDate(ByVal rangeName As String, Optional ByVal defaultValue As Variant = Empty) As Variant
    Dim rng As Range

    Set rng = GetNamedRange(rangeName)
    If rng Is Nothing Then
        GetConfigDate = defaultValue
    ElseIf IsDate(rng.Value) Then
        GetConfigDate = DateValue(CDate(rng.Value))
    Else
        GetConfigDate = defaultValue
    End If
End Function

Public Function GetConfigLong(ByVal rangeName As String, Optional ByVal defaultValue As Long = 0) As Long
    Dim rng As Range

    Set rng = GetNamedRange(rangeName)
    If rng Is Nothing Then
        GetConfigLong = defaultValue
    ElseIf IsNumeric(rng.Value) Then
        GetConfigLong = CLng(rng.Value)
    Else
        GetConfigLong = defaultValue
    End If
End Function

Public Sub SetConfigValue(ByVal rangeName As String, ByVal valueToWrite As Variant)
    Dim rng As Range

    Set rng = GetNamedRange(rangeName)
    If rng Is Nothing Then Exit Sub
    If VarType(valueToWrite) = vbString Then
        rng.NumberFormat = "@"
        rng.Value = CStr(valueToWrite)
    ElseIf IsDate(valueToWrite) Then
        rng.NumberFormat = "m/d/yyyy"
        rng.Value = CDate(valueToWrite)
    ElseIf IsEmpty(valueToWrite) Then
        rng.ClearContents
    Else
        rng.NumberFormat = "General"
        rng.Value = valueToWrite
    End If
End Sub

Public Function CommandTimeoutSeconds() As Long
    CommandTimeoutSeconds = GetConfigLong(NAME_COMMAND_TIMEOUT, 120)
    If CommandTimeoutSeconds <= 0 Then CommandTimeoutSeconds = 120
End Function

Public Function ConnectionTimeoutSeconds() As Long
    ConnectionTimeoutSeconds = GetConfigLong(NAME_CONNECTION_TIMEOUT, 15)
    If ConnectionTimeoutSeconds <= 0 Then ConnectionTimeoutSeconds = 15
End Function

Public Function ServerName() As String
    ServerName = GetConfigText(NAME_SERVER_NAME, "STL-SQL1\CRMDB")
End Function

Public Function DatabaseName() As String
    DatabaseName = GetConfigText(NAME_DATABASE_NAME, "sqlb00")
End Function

Public Function NzText(ByVal someValue As Variant, Optional ByVal defaultValue As String = "") As String
    If IsError(someValue) Or IsNull(someValue) Or IsEmpty(someValue) Then
        NzText = defaultValue
    Else
        NzText = CStr(someValue)
    End If
End Function

Public Function NzNumber(ByVal someValue As Variant, Optional ByVal defaultValue As Double = 0#) As Double
    If IsError(someValue) Or IsNull(someValue) Or IsEmpty(someValue) Or Not IsNumeric(someValue) Then
        NzNumber = defaultValue
    Else
        NzNumber = CDbl(someValue)
    End If
End Function
