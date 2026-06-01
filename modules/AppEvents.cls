VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public WithEvents App As Application

Private Sub App_SheetActivate(ByVal Sh As Object)
    RestoreAssignments Sh.Parent
    RefreshRibbon
End Sub

Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    HandleDoubleClick Sh, Target, Cancel
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    RestoreAssignments Wb
    RefreshRibbon
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    RefreshRibbon
End Sub
