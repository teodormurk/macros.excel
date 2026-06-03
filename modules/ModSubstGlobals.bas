Attribute VB_Name = "ModSubstGlobals"
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
Public Const rcCompare = 6

Public g_SheetSrc As Worksheet
Public g_SheetDst As Worksheet
Public g_Ribbon As IRibbonUI
Public g_AppEvents As AppEvents
