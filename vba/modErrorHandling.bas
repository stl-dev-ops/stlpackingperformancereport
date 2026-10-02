Attribute VB_Name = "modErrorHandling"
Option Explicit

Public Type ExcelApplicationState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
    StatusBar As Variant
End Type

Public Function CaptureApplicationState() As ExcelApplicationState
    With Application
        CaptureApplicationState.ScreenUpdating = .ScreenUpdating
        CaptureApplicationState.EnableEvents = .EnableEvents
        CaptureApplicationState.DisplayAlerts = .DisplayAlerts
        CaptureApplicationState.Calculation = .Calculation
        CaptureApplicationState.StatusBar = .StatusBar
    End With
End Function

Public Sub PrepareApplicationForRefresh()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = "Refreshing " & REPORT_TITLE & "..."
    End With
End Sub

Public Sub RestoreApplicationState(ByRef savedState As ExcelApplicationState)
    With Application
        .ScreenUpdating = savedState.ScreenUpdating
        .EnableEvents = savedState.EnableEvents
        .DisplayAlerts = savedState.DisplayAlerts
        .Calculation = savedState.Calculation
        If IsEmpty(savedState.StatusBar) Or IsNull(savedState.StatusBar) Then
            .StatusBar = False
        ElseIf VarType(savedState.StatusBar) = vbBoolean Then
            .StatusBar = False
        ElseIf Len(Trim$(CStr(savedState.StatusBar))) = 0 Then
            .StatusBar = False
        Else
            .StatusBar = savedState.StatusBar
        End If
    End With
End Sub

Public Function SanitizedErrorMessage(ByVal messageText As String) As String
    Dim sanitized As String

    sanitized = Replace(messageText, vbCr, " ")
    sanitized = Replace(sanitized, vbLf, " ")
    sanitized = Replace(sanitized, "Provider=", "")
    sanitized = Replace(sanitized, "Password=", "")
    SanitizedErrorMessage = Trim$(sanitized)
End Function

Public Function ElapsedMilliseconds(ByVal startTick As Double) As Double
    Dim elapsedSeconds As Double

    elapsedSeconds = Timer - startTick
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#
    ElapsedMilliseconds = elapsedSeconds * 1000#
End Function
