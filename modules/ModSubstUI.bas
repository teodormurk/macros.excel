Attribute VB_Name = "ModSubstUI"
Option Explicit

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set g_Ribbon = ribbon
End Sub

Public Sub getEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = False
    Select Case control.Id
        Case "btnSrcSheet", "btnDstSheet"
            returnedVal = True
        Case "btnClear"
            returnedVal = Not (g_SheetSrc Is Nothing And g_SheetDst Is Nothing)
        Case "btnColKey", "btnColSrc", "btnColDst"
            returnedVal = GetSheetRole(ActiveSheet) <> NilRole
        Case "btnReplaceAll", "btnFillEmpty", "btnMinimize", "btnMaximize", "btnAddItems", "btnSumValues"
            returnedVal = Not (g_SheetSrc Is Nothing Or g_SheetDst Is Nothing)
    End Select
End Sub

Public Sub getLabel(control As IRibbonControl, ByRef returnedVal)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim role As Long
    role = GetSheetRole(ws)
    Select Case control.Id
        Case "btnSrcSheet"
            If role = SrcRole Then
                returnedVal = "\u041b\u0438\u0441\u0442-\u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a [" & ws.Name & "]"
            Else
                returnedVal = "\u041b\u0438\u0441\u0442-\u0438\u0441\u0442\u043e\u0447\u043d\u0438\u043a"
            End If
        Case "btnDstSheet"
            If role = DstRole Then
                returnedVal = "\u041b\u0438\u0441\u0442-\u043f\u043e\u043b\u0443\u0447\u0430\u0442\u0435\u043b\u044c [" & ws.Name & "]"
            Else
                returnedVal = "\u041b\u0438\u0441\u0442-\u043f\u043e\u043b\u0443\u0447\u0430\u0442\u0435\u043b\u044c"
            End If
        Case Else
            returnedVal = control.Label
    End Select
End Sub

Public Sub btnSrcSheet_click(control As IRibbonControl)
    ToggleSheetAssignment ActiveSheet, SrcRole
End Sub

Public Sub btnDstSheet_click(control As IRibbonControl)
    ToggleSheetAssignment ActiveSheet, DstRole
End Sub

Public Sub btnClear_click(control As IRibbonControl)
    ClearAllAssignments
End Sub

Public Sub btnColKey_click(control As IRibbonControl)
    AssignSelectedColumns KeyRole
End Sub

Public Sub btnColSrc_click(control As IRibbonControl)
    AssignSelectedColumns SrcRole
End Sub

Public Sub btnColDst_click(control As IRibbonControl)
    AssignSelectedColumns DstRole
End Sub

Public Sub btnReplaceAll_click(control As IRibbonControl)
    DoReplace rcAll
End Sub

Public Sub btnFillEmpty_click(control As IRibbonControl)
    DoReplace rcEmpty
End Sub

Public Sub btnMinimize_click(control As IRibbonControl)
    DoReplace rcMin
End Sub

Public Sub btnMaximize_click(control As IRibbonControl)
    DoReplace rcMax
End Sub

Public Sub btnAddItems_click(control As IRibbonControl)
    DoReplace rcAdd
End Sub

Public Sub btnSumValues_click(control As IRibbonControl)
    DoReplace rcSum
End Sub
