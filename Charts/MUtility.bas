Attribute VB_Name = "MUtility"
Option Explicit

Sub PerformanceMode(ByVal state As Boolean)
    If state Then
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
    Else
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub

Function LastRowInOneColumn(ByVal col As String, ByRef wks As Worksheet) As Long
'Find the last used row in a Column
    Dim lastRow As Long
    With wks
        LastRowInOneColumn = .Cells(.Rows.Count, col).End(xlUp).row
    End With
End Function
