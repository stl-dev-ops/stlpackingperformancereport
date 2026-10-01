Attribute VB_Name = "modInteractionLogging"
Option Explicit

Public Function NormalizeHeaderArray(ByVal sourceHeaders As Variant) As Variant
    Dim target() As Variant
    Dim index As Long

    ReDim target(1 To UBound(sourceHeaders) - LBound(sourceHeaders) + 1)
    For index = LBound(sourceHeaders) To UBound(sourceHeaders)
        target(index - LBound(sourceHeaders) + 1) = sourceHeaders(index)
    Next index
    NormalizeHeaderArray = target
End Function

Public Function EnsureInteractionLogTable() As ListObject
    Dim headers As Variant
    Dim ws As Worksheet

    headers = NormalizeHeaderArray(Array("Timestamp", "WindowsUsername", "EventName", "Status", "Details", "SelectionContext", "WorkbookVersion"))
    Set ws = EnsureWorksheet(SHEET_INTERACTION_LOG)
    Set EnsureInteractionLogTable = EnsureListObject(ws, TABLE_INTERACTION_LOG, "A1", headers)
End Function

Public Function EnsureTestResultsTable() As ListObject
    Dim headers As Variant
    Dim ws As Worksheet

    headers = NormalizeHeaderArray(Array("RunId", "Timestamp", "Scenario", "StepName", "Status", "Details", "ArtifactPath"))
    Set ws = EnsureWorksheet(SHEET_TEST_RESULTS)
    Set EnsureTestResultsTable = EnsureListObject(ws, TABLE_TEST_RESULTS, "A1", headers)
End Function

Public Sub AppendInteractionLog(ByVal eventName As String, ByVal actionStatus As String, Optional ByVal detailText As String = "", Optional ByVal selectionContext As String = "")
    Dim lo As ListObject
    Dim newRow As ListRow

    On Error Resume Next
    Set lo = EnsureInteractionLogTable()
    Set newRow = lo.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = Now
        .Cells(1, 2).Value = Environ$("USERNAME")
        .Cells(1, 3).Value = eventName
        .Cells(1, 4).Value = actionStatus
        .Cells(1, 5).Value = SanitizedErrorMessage(detailText)
        .Cells(1, 6).Value = selectionContext
        .Cells(1, 7).Value = REPORT_VERSION
    End With
    On Error GoTo 0
End Sub

Public Function NewTestRunId(Optional ByVal scenarioHint As String = "") As String
    NewTestRunId = Format$(Now, "yyyymmdd_hhnnss") & "_" & SanitizeRunToken(scenarioHint)
End Function

Public Sub AppendTestResult(ByVal runId As String, ByVal scenarioName As String, ByVal stepName As String, ByVal passed As Boolean, Optional ByVal details As String = "", Optional ByVal artifactPath As String = "")
    Dim lo As ListObject
    Dim newRow As ListRow

    On Error Resume Next
    Set lo = EnsureTestResultsTable()
    Set newRow = lo.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = runId
        .Cells(1, 2).Value = Now
        .Cells(1, 3).Value = scenarioName
        .Cells(1, 4).Value = stepName
        .Cells(1, 5).Value = IIf(passed, "PASS", "FAIL")
        .Cells(1, 6).Value = SanitizedErrorMessage(details)
        .Cells(1, 7).Value = artifactPath
    End With
    On Error GoTo 0
End Sub

Private Function SanitizeRunToken(ByVal sourceText As String) As String
    Dim cleaned As String

    cleaned = LCase$(Trim$(sourceText))
    cleaned = Replace(cleaned, " ", "_")
    cleaned = Replace(cleaned, "/", "_")
    cleaned = Replace(cleaned, "\", "_")
    cleaned = Replace(cleaned, ":", "_")
    If Len(cleaned) = 0 Then cleaned = "run"
    SanitizeRunToken = cleaned
End Function