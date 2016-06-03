Attribute VB_Name = "MUtilities"
Option Explicit

Function LastRowInOneColumn(ByVal col As String, ByRef wks As Worksheet) As Long
'Find the last used row in a Column
    Dim lastRow As Long
    With wks
        LastRowInOneColumn = .Cells(.Rows.Count, col).End(xlUp).Row
    End With
End Function
