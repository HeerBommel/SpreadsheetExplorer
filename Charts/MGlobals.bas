Attribute VB_Name = "MGlobals"
Option Explicit

Public Const PATH As String = "X:\TUDelft\E\EUSES\ContainCharts\"
Public Const EXT As String = "*.*"
Public currentFile As String
Public outputSheets As Variant
Public chartTypes As Dictionary

Sub initGlobals()
    outputSheets = Array("ErrorLog", "ChartInfo")
    Set chartTypes = GetChartTypes()
End Sub

Function GetChartTypes() As Dictionary
    Dim chartTypes As Dictionary, wks As Worksheet, row As Integer
    Set chartTypes = New Dictionary
    Set wks = ActiveWorkbook.Sheets("ChartTypes")
    With wks
        For row = 1 To wks.UsedRange.Rows.Count
            chartTypes.Add CInt(.Cells(row, 1)), .Cells(row, 2)
        Next row
    End With
    Set GetChartTypes = chartTypes
End Function
