Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
WScript.Echo "Step 1: Excel started"

Set wb = excel.Workbooks.Open(excel.GetOpenFilename("Excel Files (*.xlsm),*.xlsm", , "Open your workbook"))
If wb Is Nothing Then
    WScript.Echo "ERROR: No file selected"
    excel.Quit
    WScript.Quit 1
End If
WScript.Echo "Step 2: Opened " & wb.Name

On Error Resume Next
Set vb = wb.VBProject
errNum = Err.Number
On Error GoTo 0

If errNum <> 0 Then
    WScript.Echo "ERROR: VBProject access denied (error " & errNum & ")"
    WScript.Echo "Enable: Excel -> Options -> Trust Center -> Macro Settings -> Trust access to VBA project"
    wb.Close False
    excel.Quit
    WScript.Quit 1
End If
WScript.Echo "Step 3: VBProject OK"

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
WScript.Echo "Step 4: Script dir = " & scriptDir

files = Array("ModSubstGlobals.bas", "AppEvents.cls", "ModSubstCore.bas", "ModSubstUtils.bas", "ModSubstUI.bas")
For i = 0 To UBound(files)
    filePath = scriptDir & "\modules\" & files(i)
    WScript.Echo "Importing: " & filePath & " ..."
    If Not fso.FileExists(filePath) Then
        WScript.Echo "  ERROR: file not found!"
    Else
        On Error Resume Next
        vb.VBComponents.Import filePath
        errNum = Err.Number
        If errNum <> 0 Then
            WScript.Echo "  ERROR: " & errNum
        Else
            WScript.Echo "  OK"
        End If
        On Error GoTo 0
    End If
Next

twPath = scriptDir & "\modules\ThisWorkbook.cls"
If fso.FileExists(twPath) Then
    WScript.Echo "Updating ThisWorkbook ..."
    Set cm = vb.VBComponents("ThisWorkbook").CodeModule
    cm.DeleteLines 1, cm.CountOfLines
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
    WScript.Echo "  OK"
End If

WScript.Echo "Saving ..."
wb.Save
WScript.Echo "Done! Modules imported successfully."
WScript.Echo ""
WScript.Echo "Press Enter to close Excel..."
WScript.StdIn.ReadLine
wb.Close
excel.Quit