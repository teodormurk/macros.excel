Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\install_log.txt"
Set logFile = fso.CreateTextFile(logPath, True, True)

Function ReadFileANSI(path)
    Set f = fso.OpenTextFile(path, 1, False, 0)
    ReadFileANSI = f.ReadAll
    f.Close
End Function

Function StripAttributes(code)
    lines = Split(code, vbCrLf)
    result = ""
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If Left(line, 10) <> "Attribute " And Left(line, 7) <> "VERSION " And line <> "BEGIN" And line <> "END" Then
            result = result & lines(i) & vbCrLf
        End If
    Next
    StripAttributes = result
End Function

Function StripClsHeader(code)
    lines = Split(code, vbCrLf)
    result = ""
    started = False
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If Not started Then
            If Left(line, 7) = "Option " Or Left(line, 7) = "Private " Or Left(line, 7) = "Public " Then
                started = True
            End If
        End If
        If started Then
            result = result & lines(i) & vbCrLf
        End If
    Next
    StripClsHeader = result
End Function

logFile.WriteLine "=== Install Log ==="

Set wb = excel.Workbooks.Open(excel.GetOpenFilename("Excel Files (*.xlsm),*.xlsm", , "Open workbook"))
If wb Is Nothing Then
    logFile.WriteLine "ERROR: No file selected"
    logFile.Close
    excel.Quit
    WScript.Quit 1
End If
logFile.WriteLine "Workbook: " & wb.Name

On Error Resume Next
Set vb = wb.VBProject
errNum = Err.Number
On Error GoTo 0

If errNum <> 0 Then
    logFile.WriteLine "ERROR: VBProject access denied"
    logFile.Close
    wb.Close False
    excel.Quit
    WScript.Quit 1
End If
logFile.WriteLine "VBProject: OK"

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
logFile.WriteLine "Script dir: " & scriptDir

basFiles = Array("ModSubstGlobals.bas", "ModSubstCore.bas", "ModSubstUtils.bas", "ModSubstUI.bas")
For i = 0 To UBound(basFiles)
    filePath = scriptDir & "\modules\" & basFiles(i)
    logFile.Write "Import " & basFiles(i) & ": "

    If Not fso.FileExists(filePath) Then
        logFile.WriteLine "FILE NOT FOUND"
    Else
        code = ReadFileANSI(filePath)
        code = StripAttributes(code)

        modName = Replace(basFiles(i), ".bas", "")
        On Error Resume Next
        vb.VBComponents.Remove vb.VBComponents(modName)
        Err.Clear

        Set comp = vb.VBComponents.Add(1)
        comp.Name = modName
        comp.CodeModule.AddFromString code
        errNum = Err.Number
        On Error GoTo 0

        If errNum <> 0 Then
            logFile.WriteLine "ERROR " & errNum
        Else
            logFile.WriteLine "OK (lines=" & comp.CodeModule.CountOfLines & ")"
        End If
    End If
Next

clsPath = scriptDir & "\modules\AppEvents.cls"
logFile.Write "Import AppEvents.cls: "
If fso.FileExists(clsPath) Then
    code = ReadFileANSI(clsPath)
    code = StripClsHeader(code)

    On Error Resume Next
    vb.VBComponents.Remove vb.VBComponents("AppEvents")
    Err.Clear

    Set cls2 = vb.VBComponents.Add(2)
    cls2.Name = "AppEvents"
    cls2.CodeModule.AddFromString code
    errNum = Err.Number
    On Error GoTo 0

    If errNum <> 0 Then
        logFile.WriteLine "ERROR " & errNum
    Else
        logFile.WriteLine "OK (lines=" & cls2.CodeModule.CountOfLines & ")"
    End If
Else
    logFile.WriteLine "FILE NOT FOUND"
End If

twPath = scriptDir & "\modules\ThisWorkbook.cls"
If fso.FileExists(twPath) Then
    logFile.Write "Update ThisWorkbook: "
    code = ReadFileANSI(twPath)
    code = StripClsHeader(code)

    On Error Resume Next
    Set cm = vb.VBComponents("ThisWorkbook").CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    Err.Clear
    cm.AddFromString code
    errNum = Err.Number
    On Error GoTo 0

    If errNum <> 0 Then
        logFile.WriteLine "ERROR " & errNum
    Else
        logFile.WriteLine "OK"
    End If
End If

logFile.WriteLine ""
logFile.WriteLine "=== All Components ==="
For Each comp In vb.VBComponents
    logFile.WriteLine "  " & comp.Name & " type=" & comp.Type & " lines=" & comp.CodeModule.CountOfLines
Next

On Error Resume Next
wb.Save
errNum = Err.Number
On Error GoTo 0
logFile.WriteLine ""
logFile.WriteLine "Save: " & errNum

logFile.WriteLine ""
logFile.WriteLine "=== After Save ==="
For Each comp In vb.VBComponents
    logFile.WriteLine "  " & comp.Name & " type=" & comp.Type & " lines=" & comp.CodeModule.CountOfLines
Next

logFile.Close
WScript.Echo "Done! See install_log.txt"
