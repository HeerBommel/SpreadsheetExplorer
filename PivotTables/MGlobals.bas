Attribute VB_Name = "MGlobals"
Option Explicit

Public Const PATH As String = "X:\TUDelft\S\SEMS16\ContainsPivot\"
Public Const EXT As String = "*.*"
Public currentFile As String
Public outputSheets As Variant
Public chartTypes As Dictionary

Sub initGlobals()
    outputSheets = Array("ErrorLog", "PivotInfo", "DataFieldInfo")
End Sub


