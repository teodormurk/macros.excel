Attribute VB_Name = "ModSubstUtils"
Option Explicit

Public Const ColorSrc As Long = 13434828
Public Const ColorDst As Long = 16764057
Public Const ColorKey As Long = 16777164
Public Const ColorNone As Long = -4142

Public Function LastUsedRow(ws As Worksheet) As Long
    On Error Resume Next
    Dim lr As Long
    lr = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If lr < 1 Then lr = 1
    LastUsedRow = lr
    On Error GoTo 0
End Function

Public Function GetKeyColumns(ws As Worksheet) As Collection
    Dim cols As New Collection
    Dim c As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If ws.Cells(1, c).Interior.Color = ColorKey Then cols.Add c
    Next c
    Set GetKeyColumns = cols
End Function

Public Function GetDataColumns(ws As Worksheet) As Collection
    Dim cols As New Collection
    Dim c As Long
    Dim lastCol As Long
    Dim roleColor As Long
    If Not g_SheetSrc Is Nothing Then
        If Not ws Is Nothing Then
            If ws Is g_SheetSrc Then roleColor = ColorSrc
        End If
    End If
    If Not g_SheetDst Is Nothing Then
        If Not ws Is Nothing Then
            If ws Is g_SheetDst Then roleColor = ColorDst
        End If
    End If
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If ws.Cells(1, c).Interior.Color = roleColor Then cols.Add c
    Next c
    Set GetDataColumns = cols
End Function

Public Function GetSheetRole(ws As Worksheet) As Long
    GetSheetRole = NilRole
    If Not g_SheetSrc Is Nothing Then
        If ws Is g_SheetSrc Then
            GetSheetRole = SrcRole
            Exit Function
        End If
    End If
    If Not g_SheetDst Is Nothing Then
        If ws Is g_SheetDst Then
            GetSheetRole = DstRole
            Exit Function
        End If
    End If
End Function

Public Sub ClearAllAssignments()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ClearSheetColors ws
    Next ws
    Set g_SheetSrc = Nothing
    Set g_SheetDst = Nothing
    RefreshRibbon
End Sub

Public Sub ClearSheetColors(ws As Worksheet)
    Dim c As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If ws.Cells(1, c).Interior.Color = ColorSrc Or _
           ws.Cells(1, c).Interior.Color = ColorDst Or _
           ws.Cells(1, c).Interior.Color = ColorKey Then
            ws.Cells(1, c).Interior.ColorIndex = ColorNone
        End If
    Next c
End Sub

Public Sub ToggleSheetAssignment(ws As Worksheet, role As Long)
    Dim current As Long
    current = GetSheetRole(ws)
    If current = role Then
        ClearSheetColors ws
        Select Case role
            Case SrcRole: Set g_SheetSrc = Nothing
            Case DstRole: Set g_SheetDst = Nothing
        End Select
    ElseIf current = NilRole Then
        Select Case role
            Case SrcRole
                If Not g_SheetSrc Is Nothing Then ClearSheetColors g_SheetSrc
                Set g_SheetSrc = ws
            Case DstRole
                If Not g_SheetDst Is Nothing Then ClearSheetColors g_SheetDst
                Set g_SheetDst = ws
        End Select
    End If
    RefreshRibbon
End Sub

Public Sub AssignSelectedColumns(role As Long)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim sheetRole As Long
    sheetRole = GetSheetRole(ws)
    If sheetRole = NilRole Then
        MsgBox "\u0421\u043d\u0430\u0447\u0430\u043b\u0430 \u043d\u0430\u0437\u043d\u0430\u0447\u044c\u0442\u0435 \u043b\u0438\u0441\u0442", vbExclamation, "\u041f\u043e\u0434\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430"
        Exit Sub
    End If
    Dim sel As Range
    Set sel = Selection
    Dim colColor As Long
    Select Case role
        Case KeyRole: colColor = ColorKey
        Case SrcRole: colColor = ColorSrc
        Case DstRole: colColor = ColorDst
    End Select
    Dim c As Range
    For Each c In sel.Columns
        ws.Cells(1, c.Column).Interior.Color = colColor
    Next c
End Sub

Public Function ValidateSetup() As String
    ValidateSetup = ""
    If g_SheetSrc Is Nothing Then
        ValidateSetup = "\u041d\u0435 \u043d\u0430\u0437\u043d\u0430\u0447\u0435\u043d \u043b\u0438\u0441\u0442-\u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a"
        Exit Function
    End If
    If g_SheetDst Is Nothing Then
        ValidateSetup = "\u041d\u0435 \u043d\u0430\u0437\u043d\u0430\u0447\u0435\u043d \u043b\u0438\u0441\u0442-\u043f\u043e\u043b\u0443\u0447\u0430\u0442\u0435\u043b\u044c"
        Exit Function
    End If
    Dim KeyColsSrc As Collection
    Dim KeyColsDst As Collection
    Set KeyColsSrc = GetKeyColumns(g_SheetSrc)
    Set KeyColsDst = GetKeyColumns(g_SheetDst)
    If KeyColsSrc.Count = 0 Or KeyColsDst.Count = 0 Then
        ValidateSetup = "\u041d\u0435 \u043d\u0430\u0437\u043d\u0430\u0447\u0435\u043d\u044b \u043a\u043b\u044e\u0447\u0435\u0432\u044b\u0435 \u0441\u0442\u043e\u043b\u0431\u0446\u044b"
        Exit Function
    End If
    Dim DataColsSrc As Collection
    Dim DataColsDst As Collection
    Set DataColsSrc = GetDataColumns(g_SheetSrc)
    Set DataColsDst = GetDataColumns(g_SheetDst)
    If DataColsSrc.Count = 0 Or DataColsDst.Count = 0 Then
        ValidateSetup = "\u041d\u0435 \u043d\u0430\u0437\u043d\u0430\u0447\u0435\u043d\u044b \u0441\u0442\u043e\u043b\u0431\u0446\u044b \u0434\u0430\u043d\u043d\u044b\u0445"
        Exit Function
    End If
End Function

Public Sub RestoreAssignments(wb As Workbook)
    If wb Is Nothing Then Exit Sub
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim srcCols As Collection
        Dim dstCols As Collection
        Dim keyCols As Collection
        Set srcCols = New Collection
        Set dstCols = New Collection
        Set keyCols = New Collection
        Dim lastCol As Long
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Dim c As Long
        For c = 1 To lastCol
            Dim clr As Long
            clr = ws.Cells(1, c).Interior.Color
            If clr = ColorSrc Then
                srcCols.Add c
            ElseIf clr = ColorDst Then
                dstCols.Add c
            ElseIf clr = ColorKey Then
                keyCols.Add c
            End If
        Next c
        If srcCols.Count > 0 Then Set g_SheetSrc = ws
        If dstCols.Count > 0 Then Set g_SheetDst = ws
    Next ws
End Sub

Public Sub RefreshRibbon()
    If Not g_Ribbon Is Nothing Then
        g_Ribbon.Invalidate
    End If
End Sub

Public Sub HandleDoubleClick(Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim ws As Worksheet
    Set ws = Sh
    Dim sheetRole As Long
    sheetRole = GetSheetRole(ws)
    If sheetRole = NilRole Then Exit Sub
    Dim col As Long
    col = Target.Column
    Dim currentColor As Long
    currentColor = ws.Cells(1, col).Interior.Color
    If currentColor = ColorSrc Or currentColor = ColorDst Or currentColor = ColorKey Then
        ws.Cells(1, col).Interior.ColorIndex = ColorNone
        Cancel = True
        Exit Sub
    End If
    Dim response As VbMsgBoxResult
    If sheetRole = SrcRole Then
        response = MsgBox("\u041f\u043e\u043c\u0435\u0442\u0438\u0442\u044c \u043a\u0430\u043a \u0441\u0442\u043e\u043b\u0431\u0435\u0446-\u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a?", vbYesNoCancel + vbQuestion, "\u041f\u043e\u0434\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430")
        If response = vbYes Then
            ws.Cells(1, col).Interior.Color = ColorSrc
        ElseIf response = vbNo Then
            response = MsgBox("\u041f\u043e\u043c\u0435\u0442\u0438\u0442\u044c \u043a\u0430\u043a \u043a\u043b\u044e\u0447?", vbYesNo + vbQuestion, "\u041f\u043e\u0434\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430")
            If response = vbYes Then
                ws.Cells(1, col).Interior.Color = ColorKey
            End If
        End If
    ElseIf sheetRole = DstRole Then
        response = MsgBox("\u041f\u043e\u043c\u0435\u0442\u0438\u0442\u044c \u043a\u0430\u043a \u0441\u0442\u043e\u043b\u0431\u0435\u0446-\u043f\u043e\u043b\u0443\u0447\u0430\u0442\u0435\u043b\u044c?", vbYesNoCancel + vbQuestion, "\u041f\u043e\u0434\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430")
        If response = vbYes Then
            ws.Cells(1, col).Interior.Color = ColorDst
        ElseIf response = vbNo Then
            response = MsgBox("\u041f\u043e\u043c\u0435\u0442\u0438\u0442\u044c \u043a\u0430\u043a \u043a\u043b\u044e\u0447?", vbYesNo + vbQuestion, "\u041f\u043e\u0434\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430")
            If response = vbYes Then
                ws.Cells(1, col).Interior.Color = ColorKey
            End If
        End If
    End If
    Cancel = True
End Sub
