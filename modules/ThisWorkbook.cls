VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Set g_AppEvents = New AppEvents
    Set g_AppEvents.App = Application
    RefreshRibbon
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Set g_AppEvents = Nothing
End Sub
