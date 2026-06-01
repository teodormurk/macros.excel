Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\install_log.txt"
Set logFile = fso.CreateTextFile(logPath, True, True)

Function ReadFile1251(path)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "windows-1251"
    stream.Open
    stream.LoadFromFile path
    ReadFile1251 = stream.ReadText
    stream.Close
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
        On Error Resume Next
        code = ReadFile1251(filePath)
        errNum = Err.Number
        On Error GoTo 0

        If errNum <> 0 Then
            logFile.WriteLine "Read ERROR " & errNum & " (ADODB not available, using FSO)"
            Set f2 = fso.OpenTextFile(filePath, 1, False, -1)
            code = f2.ReadAll
            f2.Close
        End If

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
    On Error Resume Next
    code = ReadFile1251(clsPath)
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then
        Set f2 = fso.OpenTextFile(clsPath, 1, False, -1)
        code = f2.ReadAll
        f2.Close
    End If

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
    On Error Resume Next
    code = ReadFile1251(twPath)
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then
        Set f2 = fso.OpenTextFile(twPath, 1, False, -1)
        code = f2.ReadAll
        f2.Close
    End If

    Set cm = vb.VBComponents("ThisWorkbook").CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    Err.Clear

    lines = Split(code, vbCrLf)
    started = False
    lineNum = 0
    For li = 0 To UBound(lines)
        line = lines(li)
        If Not started Then
            If Left(Trim(line), 10) = "Attribute " Then started = True
        End If
        If started Then
            lineNum = lineNum + 1
            cm.InsertLines lineNum, line
        End If
    Next
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
