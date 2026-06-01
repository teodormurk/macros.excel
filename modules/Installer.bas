Attribute VB_Name = "Installer"
Option Explicit

Public Sub InstallModules()
    Dim basePath As String
    basePath = ThisWorkbook.Path

    If basePath = "" Then
        MsgBox "Сначачите не удалось сохранить книги. " & _
               "Сохрните и запустите update снова.", vbExclamation, "Error"
        Exit Sub
    End If

    Dim modulesPath As String
    modulesPath = basePath & "Модули" & Application.PathSeparator & "modules"

    If Dir(modulesPath, vbDirectory) = "" Then
        MsgBox "Папка modules/ не найдена: " & modulesPath & vbCrLf & vbCrLf & _
               "Положите книгу " & Chr(34) & basePath & Chr(34) & " рядом с update.exe", vbExclamation, "Error"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("ModSubstCore")
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("ModSubstUtils")
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("ModSubstUI")
    On Error GoTo 0

    Dim f As String
    f = Dir(modulesPath & Application.PathSeparator & "*.bas")
    Do While f <> ""
        If f <> "Installer.bas" Then
            Dim importPath As String
            importPath = modulesPath & Application.PathSeparator & f
            If Dir(importPath) <> "" Then
                ThisWorkbook.VBProject.VBComponents.Import importPath
            End If
        End If
        f = Dir()
    Loop

    Dim twPath As String
    twPath = modulesPath & Application.PathSeparator & "ThisWorkbook.cls"
    If Dir(twPath) <> "" Then
        UpdateThisWorkbook twPath
    End If

    ThisWorkbook.Save
    MsgBox "Modules updated!", vbInformation, "Done"
End Sub

Private Sub UpdateThisWorkbook(filePath As String)
    Dim cm As CodeModule
    Set cm = ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
    cm.DeleteLines 1, cm.CountOfLines

    Dim fnum As Integer
    fnum = FreeFile
    Open filePath For Input As fnum
    Dim line As String
    Dim lineNum As Long
    lineNum = 0
    Dim started As Boolean
    started = False
    Do Until EOF(fnum)
        Line Input #fnum, line
        If Not started Then
            If Len(Trim(line)) > 0 Then
                If Left(Trim(line), 7) = "Option " Then
                    started = True
                End If
            End If
        End If
        If started Then
            lineNum = lineNum + 1
            cm.InsertLines lineNum, line
        End If
    Loop
    Close fnum
End Sub
