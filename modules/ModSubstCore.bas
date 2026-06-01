Attribute VB_Name = "ModSubstCore"
Option Explicit

Public Const NilRole = 0
Public Const SrcRole = 1
Public Const DstRole = 2
Public Const KeyRole = 3

Public Const rcAll = 0
Public Const rcEmpty = 1
Public Const rcMin = 2
Public Const rcMax = 3
Public Const rcAdd = 4
Public Const rcSum = 5

Public g_SheetSrc As Worksheet
Public g_SheetDst As Worksheet
Public g_Ribbon As IRibbonUI
Public g_AppEvents As AppEvents

Public Function NormalizeKey(ByVal Value As String) As String
    NormalizeKey = UCase(Replace(Trim(Value), " ", ""))
End Function

Public Function BuildCompositeKey(ws As Worksheet, keyCols As Collection, rowNum As Long) As String
    Dim result As String
    Dim i As Long
    result = ""
    For i = 1 To keyCols.Count
        Dim cellVal As String
        cellVal = CStr(ws.Cells(rowNum, keyCols(i)).Value)
        If InStr(1, cellVal, "#", vbTextCompare) > 0 Then
            BuildCompositeKey = ""
            Exit Function
        End If
        result = result & NormalizeKey(cellVal)
    Next i
    BuildCompositeKey = result
End Function

Public Function BuildIndex() As Object
    Dim index As Object
    Set index = CreateObject("Scripting.Dictionary")
    index.CompareMode = vbTextCompare
    
    Dim keyCols As Collection
    Set keyCols = GetKeyColumns(g_SheetDst)
    
    If keyCols.Count = 0 Then
        Set BuildIndex = index
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = LastUsedRow(g_SheetDst)
    
    Dim r As Long
    For r = 2 To lastRow
        Dim compKey As String
        compKey = BuildCompositeKey(g_SheetDst, keyCols, r)
        If compKey <> "" Then
            If Not index.Exists(compKey) Then
                index.Add compKey, r
            End If
        End If
    Next r
    
    Set BuildIndex = index
End Function

Public Sub DoReplace(mode As Long)
    Dim err As String
    err = ValidateSetup()
    If err <> "" Then
        MsgBox err, vbExclamation, "�����������"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim KeyColsSrc As Collection
    Dim KeyColsDst As Collection
    Dim DataColsSrc As Collection
    Dim DataColsDst As Collection
    
    Set KeyColsSrc = GetKeyColumns(g_SheetSrc)
    Set KeyColsDst = GetKeyColumns(g_SheetDst)
    Set DataColsSrc = GetDataColumns(g_SheetSrc)
    Set DataColsDst = GetDataColumns(g_SheetDst)
    
    Dim index As Object
    Set index = BuildIndex()
    
    Dim lastRowSrc As Long
    lastRowSrc = LastUsedRow(g_SheetSrc)
    
    Dim changedCount As Long
    changedCount = 0
    Dim addedCount As Long
    addedCount = 0
    
    Dim r As Long
    For r = 2 To lastRowSrc
        Dim compKey As String
        compKey = BuildCompositeKey(g_SheetSrc, KeyColsSrc, r)
        If compKey <> "" Then
            If index.Exists(compKey) Then
                Dim dstRow As Long
                dstRow = index(compKey)
                
                Dim dstKey As String
                dstKey = BuildCompositeKey(g_SheetDst, KeyColsDst, dstRow)
                
                If compKey = dstKey Then
                    Dim i As Long
                    For i = 1 To DataColsSrc.Count
                        If i <= DataColsDst.Count Then
                            Dim srcCol As Long
                            Dim dstCol As Long
                            srcCol = DataColsSrc(i)
                            dstCol = DataColsDst(i)
                            
                            If Not g_SheetDst.Cells(dstRow, dstCol).HasFormula Then
                                Dim srcVal As Variant
                                Dim dstVal As Variant
                                srcVal = g_SheetSrc.Cells(r, srcCol).Value
                                dstVal = g_SheetDst.Cells(dstRow, dstCol).Value
                                
                                If Not IsSourceEmpty(srcVal) Then
                                    Dim result As Variant
                                    result = ApplyReplace(srcVal, dstVal, mode)
                                    If Not IsNull(result) Then
                                        g_SheetDst.Cells(dstRow, dstCol).Value = result
                                        changedCount = changedCount + 1
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If
            ElseIf mode = rcAdd Then
                AddRow r, KeyColsSrc, KeyColsDst, DataColsSrc, DataColsDst
                addedCount = addedCount + 1
            End If
        End If
    Next r
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Dim msg As String
    msg = "�������� �����: " & changedCount
    If addedCount > 0 Then
        msg = msg & "��������� �����: " & addedCount
    End If
    MsgBox msg, vbInformation, "�����������"
End Sub

Private Function IsSourceEmpty(val As Variant) As Boolean
    IsSourceEmpty = False
    If IsEmpty(val) Then
        IsSourceEmpty = True
        Exit Function
    End If
    If VarType(val) = vbString Then
        If Trim(CStr(val)) = "" Then IsSourceEmpty = True
    End If
End Function

Private Function IsDestEmpty(val As Variant) As Boolean
    IsDestEmpty = False
    If IsEmpty(val) Then
        IsDestEmpty = True
        Exit Function
    End If
    If VarType(val) = vbString Then
        If Trim(CStr(val)) = "" Then IsDestEmpty = True
    End If
End Function

Private Function ApplyReplace(srcVal As Variant, dstVal As Variant, mode As Long) As Variant
    ApplyReplace = Null
    
    Select Case mode
        Case rcAll
            ApplyReplace = srcVal
            
        Case rcEmpty
            If IsDestEmpty(dstVal) Then
                ApplyReplace = srcVal
            End If
            
        Case rcMin
            If IsNumeric(srcVal) Then
                If IsDestEmpty(dstVal) Or Not IsNumeric(dstVal) Then
                    ApplyReplace = srcVal
                ElseIf CDbl(srcVal) < CDbl(dstVal) Then
                    ApplyReplace = srcVal
                End If
            End If
            
        Case rcMax
            If IsNumeric(srcVal) Then
                If IsDestEmpty(dstVal) Or Not IsNumeric(dstVal) Then
                    ApplyReplace = srcVal
                ElseIf CDbl(srcVal) > CDbl(dstVal) Then
                    ApplyReplace = srcVal
                End If
            End If
            
        Case rcSum
            If IsNumeric(srcVal) Then
                If IsDestEmpty(dstVal) Or Not IsNumeric(dstVal) Then
                    ApplyReplace = CDbl(srcVal)
                Else
                    ApplyReplace = CDbl(srcVal) + CDbl(dstVal)
                End If
            End If
            
        Case rcAdd
            ApplyReplace = Null
    End Select
End Function

Private Sub AddRow(srcRow As Long, KeyColsSrc As Collection, KeyColsDst As Collection, DataColsSrc As Collection, DataColsDst As Collection)
    Dim lastRow As Long
    lastRow = LastUsedRow(g_SheetDst) + 1
    
    Dim i As Long
    For i = 1 To KeyColsSrc.Count
        If i <= KeyColsDst.Count Then
            g_SheetDst.Cells(lastRow, KeyColsDst(i)).Value = _
                g_SheetSrc.Cells(srcRow, KeyColsSrc(i)).Value
        End If
    Next i
    
    For i = 1 To DataColsSrc.Count
        If i <= DataColsDst.Count Then
            g_SheetDst.Cells(lastRow, DataColsDst(i)).Value = _
                g_SheetSrc.Cells(srcRow, DataColsSrc(i)).Value
        End If
    Next i
End Sub
