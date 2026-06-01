Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\install_log.txt"
Set logFile = fso.CreateTextFile(logPath, True, True)

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
        Set f = fso.OpenTextFile(filePath, 1)
        code = f.ReadAll
        f.Close

        modName = Replace(basFiles(i), ".bas", "")
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
    vb.VBComponents.Remove vb.VBComponents("AppEvents")
    Err.Clear

    Set f = fso.OpenTextFile(clsPath, 1)
    code = f.ReadAll
    f.Close

    Set cls = vb.VBComponents.Add(2)
    cls.Name = "AppEvents"
    cls.CodeModule.AddFromString code
    errNum = Err.Number
    On Error GoTo 0

    If errNum <> 0 Then
        logFile.WriteLine "ERROR " & errNum
    Else
        logFile.WriteLine "OK (lines=" & cls.CodeModule.CountOfLines & ")"
    End If
Else
    logFile.WriteLine "FILE NOT FOUND"
End If

twPath = scriptDir & "\modules\ThisWorkbook.cls"
If fso.FileExists(twPath) Then
    logFile.Write "Update ThisWorkbook: "
    On Error Resume Next
    Set cm = vb.VBComponents("ThisWorkbook").CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    Err.Clear

    Set f = fso.OpenTextFile(twPath, 1)
    code = f.ReadAll
    f.Close

    started = False
    outCode = ""
    lines = Split(code, vbCrLf)
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
