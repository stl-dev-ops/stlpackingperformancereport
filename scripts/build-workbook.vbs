On Error Resume Next

Dim xl
Dim wb
Dim vbProj
Dim comp
Dim fso
Dim shell
Dim folder
Dim file
Dim codeText
Dim outPath
Dim logPath
Dim artifactsRoot
Dim dialogWatchCommand

outPath = "C:\dev\STLPackingPerformanceReport\STLPackingPerformanceReport.xlsm"
If WScript.Arguments.Count > 0 Then outPath = WScript.Arguments(0)
folder = "C:\dev\STLPackingPerformanceReport\vba"
logPath = "C:\dev\STLPackingPerformanceReport\build-workbook.log"
artifactsRoot = "C:\dev\STLPackingPerformanceReport\artifacts"
dialogWatchCommand = "powershell -NoProfile -ExecutionPolicy Bypass -File ""C:\dev\STLPackingPerformanceReport\scripts\watch-excel-dialogs.ps1"" -OutputPath ""C:\dev\STLPackingPerformanceReport\artifacts\dialog-watch-build.log"" -TimeoutSeconds 180"

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
EnsureFolder artifactsRoot
WriteLog "Start"
shell.Run dialogWatchCommand, 0, False

Set xl = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WriteLog "ERR_EXCEL:" & Err.Number & ":" & Err.Description
    WScript.Echo "ERR_EXCEL:" & Err.Number & ":" & Err.Description
    WScript.Quit 1
End If

xl.Visible = False
xl.DisplayAlerts = False

Set wb = xl.Workbooks.Add()
If Err.Number <> 0 Then
    WriteLog "ERR_WORKBOOK:" & Err.Number & ":" & Err.Description
    WScript.Echo "ERR_WORKBOOK:" & Err.Number & ":" & Err.Description
    xl.Quit
    WScript.Quit 1
End If

Set vbProj = wb.VBProject
If Err.Number <> 0 Then
    WriteLog "ERR_VBPROJECT:" & Err.Number & ":" & Err.Description
    WScript.Echo "ERR_VBPROJECT:" & Err.Number & ":" & Err.Description
    wb.Close False
    xl.Quit
    WScript.Quit 1
End If

For Each file In fso.GetFolder(folder).Files
    If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
        WriteLog "Importing " & file.Name
        vbProj.VBComponents.Import file.Path
        If Err.Number <> 0 Then
            WriteLog "ERR_IMPORT:" & file.Name & ":" & Err.Number & ":" & Err.Description
            WScript.Echo "ERR_IMPORT:" & file.Name & ":" & Err.Number & ":" & Err.Description
            wb.Close False
            xl.Quit
            WScript.Quit 1
        End If
    End If
Next

Err.Clear
codeText = BuildEmbeddedSqlCode(fso.BuildPath(fso.GetParentFolderName(folder), "packing_estimated_vs_actual.sql"))
If Err.Number = 0 Then
    Set comp = vbProj.VBComponents.Add(1)
    comp.Name = "modEmbeddedSql"
    comp.CodeModule.AddFromString codeText
End If
If Err.Number <> 0 Then
    WriteLog "ERR_EMBED_SQL:" & Err.Number & ":" & Err.Description
    WScript.Echo "ERR_EMBED_SQL:" & Err.Number & ":" & Err.Description
    wb.Close False
    xl.Quit
    WScript.Quit 1
End If
WriteLog "Embedded packing SQL"

codeText = ReadCodeWithoutAttributes(fso.BuildPath(folder, "ThisWorkbook.cls"))
vbProj.VBComponents("ThisWorkbook").CodeModule.DeleteLines 1, vbProj.VBComponents("ThisWorkbook").CodeModule.CountOfLines
vbProj.VBComponents("ThisWorkbook").CodeModule.AddFromString codeText
If Err.Number <> 0 Then
    WriteLog "ERR_THISWORKBOOK:" & Err.Number & ":" & Err.Description
    WScript.Echo "ERR_THISWORKBOOK:" & Err.Number & ":" & Err.Description
    wb.Close False
    xl.Quit
    WScript.Quit 1
End If

Set comp = vbProj.VBComponents.Add(1)
comp.Name = "BuildProbe"
comp.CodeModule.AddFromString BuildProbeCode()
If Err.Number <> 0 Then
    WriteLog "ERR_PROBE:" & Err.Number & ":" & Err.Description
    WScript.Echo "ERR_PROBE:" & Err.Number & ":" & Err.Description
    wb.Close False
    xl.Quit
    WScript.Quit 1
End If

xl.Run wb.Name & "!BuildProbe.RunInitializer"
If CStr(wb.Worksheets(1).Range("A1").Value) <> "OK" Then
    WriteLog "ERR_INIT:" & CStr(wb.Worksheets(1).Range("A1").Value)
    WScript.Echo "ERR_INIT:" & CStr(wb.Worksheets(1).Range("A1").Value)
    wb.Close False
    xl.Quit
    WScript.Quit 1
End If

DeleteSheetIfPresent wb, "Sheet1"
On Error Resume Next
wb.Worksheets("Dashboard").Activate
On Error GoTo 0
wb.SaveAs outPath, 52
If Err.Number <> 0 Then
    WriteLog "ERR_SAVE:" & Err.Number & ":" & Err.Description
    WScript.Echo "ERR_SAVE:" & Err.Number & ":" & Err.Description
    wb.Close False
    xl.Quit
    WScript.Quit 1
End If

WriteLog "Saved workbook"
WScript.Echo "OK:" & outPath

wb.Close False
xl.Quit
WriteLog "Done"

Function BuildEmbeddedSqlCode(filePath)
    Dim ts
    Dim output

    Set ts = fso.OpenTextFile(filePath, 1)
    output = "Option Explicit" & vbCrLf & _
        "Public Function PackingSqlText() As String" & vbCrLf & _
        "    Dim sqlText As String" & vbCrLf
    Do Until ts.AtEndOfStream
        output = output & "    sqlText = sqlText & """ & Replace(ts.ReadLine, """", """""") & """ & vbCrLf" & vbCrLf
    Loop
    ts.Close
    BuildEmbeddedSqlCode = output & "    PackingSqlText = sqlText" & vbCrLf & "End Function"
End Function

Function ReadCodeWithoutAttributes(filePath)
    Dim ts
    Dim content
    Dim lines
    Dim i
    Dim output

    Set ts = fso.OpenTextFile(filePath, 1)
    content = ts.ReadAll
    ts.Close

    lines = Split(content, vbCrLf)
    output = ""

    For i = 0 To UBound(lines)
        If Not IsMetadataLine(lines(i)) Then
            If output = "" Then
                output = lines(i)
            Else
                output = output & vbCrLf & lines(i)
            End If
        End If
    Next

    ReadCodeWithoutAttributes = output
End Function

Function IsMetadataLine(lineText)
    Dim trimmed

    trimmed = Trim(lineText)
    If Left(trimmed, 7) = "VERSION" Then
        IsMetadataLine = True
    ElseIf Left(trimmed, 9) = "Attribute" Then
        IsMetadataLine = True
    ElseIf trimmed = "BEGIN" Or trimmed = "END" Then
        IsMetadataLine = True
    ElseIf Left(trimmed, 9) = "MultiUse " Then
        IsMetadataLine = True
    Else
        IsMetadataLine = False
    End If
End Function

Function BuildProbeCode()
    BuildProbeCode = _
        "Public Sub RunInitializer()" & vbCrLf & _
        "    On Error GoTo EH" & vbCrLf & _
        "    InitializeReportWorkbook" & vbCrLf & _
        "    Worksheets(1).Range(""A1"").Value = ""OK""" & vbCrLf & _
        "    Exit Sub" & vbCrLf & _
        "EH:" & vbCrLf & _
        "    Worksheets(1).Range(""A1"").Value = CStr(Err.Number) & ""::"" & Err.Description" & vbCrLf & _
        "End Sub"
End Function

Sub DeleteSheetIfPresent(workbookObj, sheetName)
    Dim ws
    On Error Resume Next
    Set ws = workbookObj.Worksheets(sheetName)
    If Err.Number = 0 Then ws.Delete
    Err.Clear
    On Error GoTo 0
End Sub

Sub EnsureFolder(folderPath)
    On Error Resume Next
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Sub WriteLog(messageText)
    Dim ts
    Set ts = fso.OpenTextFile(logPath, 8, True)
    ts.WriteLine Now & " | " & messageText
    ts.Close
End Sub
