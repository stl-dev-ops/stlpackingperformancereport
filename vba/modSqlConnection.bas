Attribute VB_Name = "modSqlConnection"
Option Explicit

Private Const adCmdText As Long = 1
Private Const adParamInput As Long = 1
Private Const adVarWChar As Long = 202

Public Function OpenCermConnection() As Object
    Dim cn As Object

    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = ConnectionTimeoutSeconds()
    cn.CommandTimeout = CommandTimeoutSeconds()
    cn.Open "Provider=SQLOLEDB;Data Source=" & ServerName() & ";Initial Catalog=" & DatabaseName() & ";Integrated Security=SSPI;"
    Set OpenCermConnection = cn
End Function

Public Function CreateTextCommand(ByVal cn As Object, ByVal sqlText As String) As Object
    Dim cmd As Object

    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandTimeout = CommandTimeoutSeconds()
        .CommandText = sqlText
    End With
    Set CreateTextCommand = cmd
End Function

Public Sub AppendTextParameter(ByVal cmd As Object, ByVal paramName As String, ByVal size As Long, ByVal paramValue As String)
    cmd.Parameters.Append cmd.CreateParameter(paramName, adVarWChar, adParamInput, size, paramValue)
End Sub

Public Function ExecuteCommandMatrix(ByVal cmd As Object, ByRef headers As Variant) As Variant
    Dim rs As Object

    Set rs = cmd.Execute
    ExecuteCommandMatrix = RecordsetToMatrix(rs, headers)

    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    Set rs = Nothing
End Function

Public Function RecordsetToMatrix(ByVal rs As Object, ByRef headers As Variant) As Variant
    Dim fieldCount As Long
    Dim rowCount As Long
    Dim sourceData As Variant
    Dim matrix() As Variant
    Dim rowIndex As Long
    Dim colIndex As Long

    fieldCount = rs.Fields.Count
    ReDim headers(1 To fieldCount)
    For colIndex = 1 To fieldCount
        headers(colIndex) = rs.Fields(colIndex - 1).Name
    Next colIndex

    If rs.EOF Then
        RecordsetToMatrix = Empty
        Exit Function
    End If

    sourceData = rs.GetRows
    rowCount = UBound(sourceData, 2) - LBound(sourceData, 2) + 1
    ReDim matrix(1 To rowCount, 1 To fieldCount)

    For rowIndex = 1 To rowCount
        For colIndex = 1 To fieldCount
            matrix(rowIndex, colIndex) = sourceData(colIndex - 1, rowIndex - 1)
        Next colIndex
    Next rowIndex

    RecordsetToMatrix = matrix
End Function

Public Function MatrixRowCount(ByVal matrix As Variant) As Long
    On Error GoTo EH
    MatrixRowCount = UBound(matrix, 1) - LBound(matrix, 1) + 1
    Exit Function
EH:
    MatrixRowCount = 0
End Function

Public Function HeaderCount(ByVal headers As Variant) As Long
    On Error GoTo EH
    HeaderCount = UBound(headers) - LBound(headers) + 1
    Exit Function
EH:
    HeaderCount = 0
End Function

Public Function LoadPackingEstimatedVsActualData(ByVal cn As Object, ByRef headers As Variant) As Variant
    Dim cmd As Object
    Dim sqlText As String

    sqlText = modEmbeddedSql.PackingSqlText()
    sqlText = ApplyConfigFilters(sqlText)
    Set cmd = CreateTextCommand(cn, sqlText)
    LoadPackingEstimatedVsActualData = ExecuteCommandMatrix(cmd, headers)
End Function

Private Function ApplyConfigFilters(ByVal sqlText As String) As String
    sqlText = ReplaceDeclareAssignment(sqlText, "@BestDateFrom", SqlDateLiteral(GetConfigDate(NAME_ESTIMATE_DATE_FROM)))
    sqlText = ReplaceDeclareAssignment(sqlText, "@BestDateTo", SqlDateLiteral(GetConfigDate(NAME_ESTIMATE_DATE_TO)))
    sqlText = ReplaceDeclareAssignment(sqlText, "@DeliveryDateFrom", SqlDateLiteral(GetConfigDate(NAME_DELIVERY_DATE_FROM)))
    sqlText = ReplaceDeclareAssignment(sqlText, "@DeliveryDateTo", SqlDateLiteral(GetConfigDate(NAME_DELIVERY_DATE_TO)))
    sqlText = ReplaceDeclareAssignment(sqlText, "@ActualWorkDateFrom", SqlDateLiteral(GetConfigDate(NAME_ACTUAL_WORK_DATE_FROM)))
    sqlText = ReplaceDeclareAssignment(sqlText, "@ActualWorkDateTo", SqlDateLiteral(GetConfigDate(NAME_ACTUAL_WORK_DATE_TO)))
    sqlText = ReplaceDeclareAssignment(sqlText, "@CustomerLike", SqlNVarCharLiteral(GetConfigText(NAME_CUSTOMER_LIKE)))
    sqlText = ReplaceDeclareAssignment(sqlText, "@WorkCenterLike", SqlNVarCharLiteral(GetConfigText(NAME_WORK_CENTER_LIKE)))
    sqlText = ReplaceDeclareAssignment(sqlText, "@EmployeeLike", SqlNVarCharLiteral(GetConfigText(NAME_EMPLOYEE_LIKE)))
    ApplyConfigFilters = sqlText
End Function

Private Function ReplaceDeclareAssignment(ByVal sqlText As String, ByVal variableName As String, ByVal replacementLiteral As String) As String
    Dim startPos As Long
    Dim lineEnd As Long
    Dim assignmentPos As Long
    Dim lineText As String
    Dim newLine As String

    startPos = InStr(1, sqlText, "DECLARE " & variableName, vbTextCompare)
    If startPos = 0 Then
        ReplaceDeclareAssignment = sqlText
        Exit Function
    End If

    lineEnd = InStr(startPos, sqlText, vbLf)
    If lineEnd = 0 Then lineEnd = Len(sqlText) + 1
    lineText = Mid$(sqlText, startPos, lineEnd - startPos)
    assignmentPos = InStr(1, lineText, "=", vbTextCompare)
    If assignmentPos = 0 Then
        ReplaceDeclareAssignment = sqlText
        Exit Function
    End If

    newLine = Left$(lineText, assignmentPos) & " " & replacementLiteral & ";"
    ReplaceDeclareAssignment = Left$(sqlText, startPos - 1) & newLine & Mid$(sqlText, lineEnd)
End Function

Public Function SqlDateLiteral(ByVal configValue As Variant) As String
    If IsEmpty(configValue) Then
        SqlDateLiteral = "NULL"
    Else
        SqlDateLiteral = "'" & Format$(CDate(configValue), "yyyy-mm-dd") & "'"
    End If
End Function

Public Function SqlNVarCharLiteral(ByVal configValue As String) As String
    If Len(Trim$(configValue)) = 0 Then
        SqlNVarCharLiteral = "NULL"
    Else
        SqlNVarCharLiteral = "N'" & Replace(configValue, "'", "''") & "'"
    End If
End Function
