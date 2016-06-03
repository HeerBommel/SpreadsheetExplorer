Attribute VB_Name = "MOutput"
Option Explicit

Sub ClearOutputSheet()
    Dim i
    For i = 0 To UBound(outputSheets)
       ActiveWorkbook.Sheets(outputSheets(i)).UsedRange.Offset(1).ClearContents
    Next i
End Sub

Sub OutputErrors(ByRef errorInfo As ErrorLog)
    Dim wks As Worksheet, lastRow As Long
    Set wks = ThisWorkbook.Sheets("ErrorLog")
    lastRow = LastRowInOneColumn("A", wks)
    wks.Cells(lastRow + 1, 1) = errorInfo.ErrorCode
    wks.Cells(lastRow + 1, 2) = errorInfo.info
End Sub

Sub OutputChartInfo(ByRef ChartInfo As ChartInfo)
    Dim wks As Worksheet, lastRow As Long
    Set wks = ThisWorkbook.Sheets("ChartInfo")
    lastRow = LastRowInOneColumn("A", wks)
    With wks
        .Cells(lastRow + 1, 1) = ChartInfo.FileName
        .Cells(lastRow + 1, 2) = ChartInfo.Emmbedded
        .Cells(lastRow + 1, 3) = ChartInfo.Worksheet
        .Cells(lastRow + 1, 4) = ChartInfo.Index
        .Cells(lastRow + 1, 5) = ChartInfo.Name
        .Cells(lastRow + 1, 6) = ChartInfo.Title
        .Cells(lastRow + 1, 7) = ChartInfo.ChartType
    End With
End Sub
