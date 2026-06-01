Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\install_log.txt"
Set logFile = fso.CreateTextFile(logPath, True, True)

logFile.WriteLine "=== Install Log ==="
logFile.WriteLine ""

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
    logFile.WriteLine "ERROR: VBProject access denied (" & errNum & ")"
    logFile.Close
    wb.Close False
    excel.Quit
    WScript.Quit 1
End If
logFile.WriteLine "VBProject: OK"

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
logFile.WriteLine "Script dir: " & scriptDir
logFile.WriteLine ""

files = Array("ModSubstGlobals.bas", "AppEvents.cls", "ModSubstCore.bas", "ModSubstUtils.bas", "ModSubstUI.bas")
For i = 0 To UBound(files)
    filePath = scriptDir & "\modules\" & files(i)
    logFile.Write "Import " & files(i) & ": "
    If Not fso.FileExists(filePath) Then
        logFile.WriteLine "FILE NOT FOUND"
    Else
        On Error Resume Next
        vb.VBComponents.Import filePath
        errNum = Err.Number
        On Error GoTo 0
        If errNum <> 0 Then
            logFile.WriteLine "ERROR " & errNum
        Else
            logFile.WriteLine "OK"
        End If
    End If
Next

logFile.WriteLine ""
logFile.WriteLine "=== Components after import ==="
For Each comp In vb.VBComponents
    logFile.WriteLine "  " & comp.Name & " (" & comp.Type & ") lines=" & comp.CodeModule.CountOfLines
Next

twPath = scriptDir & "\modules\ThisWorkbook.cls"
If fso.FileExists(twPath) Then
    logFile.Write "Updating ThisWorkbook: "
    On Error Resume Next
    Set cm = vb.VBComponents("ThisWorkbook").CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then
        logFile.WriteLine "DeleteLines ERROR " & errNum
    Else
        Set tf = fso.OpenTextFile(twPath, 1)
        lineNum = 0
        started = False
        Do While Not tf.AtEndOfStream
            line = tf.ReadLine
            If Not started Then
                If Left(Trim(line), 10) = "Attribute " Then started = True
            End If
            If started Then
                lineNum = lineNum + 1
                cm.InsertLines lineNum, line
            End If
        Loop
        tf.Close
        logFile.WriteLine "OK (" & lineNum & " lines)"
    End If
End If

logFile.WriteLine ""
logFile.WriteLine "=== Components before save ==="
For Each comp In vb.VBComponents
    logFile.WriteLine "  " & comp.Name & " (" & comp.Type & ") lines=" & comp.CodeModule.CountOfLines
Next

On Error Resume Next
wb.Save
errNum = Err.Number
On Error GoTo 0
logFile.WriteLine ""
logFile.WriteLine "Save: " & errNum

logFile.WriteLine ""
logFile.WriteLine "=== Components after save ==="
For Each comp In vb.VBComponents
    logFile.WriteLine "  " & comp.Name & " (" & comp.Type & ") lines=" & comp.CodeModule.CountOfLines
Next

logFile.Close
WScript.Echo "Done! See install_log.txt"
