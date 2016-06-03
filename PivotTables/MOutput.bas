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

Sub OutputPivotInfo(ByRef ptInfo As PivotInfo)
    Dim wks As Worksheet, lastRow As Long
    Set wks = ThisWorkbook.Sheets("PivotInfo")
    lastRow = LastRowInOneColumn("A", wks)
    With wks
        .Cells(lastRow + 1, 1) = ptInfo.Workbook
        .Cells(lastRow + 1, 2) = ptInfo.Worksheet
        .Cells(lastRow + 1, 3) = ptInfo.Name
        .Cells(lastRow + 1, 4) = ptInfo.memory
        .Cells(lastRow + 1, 5) = ptInfo.records
        .Cells(lastRow + 1, 6) = ptInfo.DataFields
        .Cells(lastRow + 1, 7) = ptInfo.RowFields
        .Cells(lastRow + 1, 8) = ptInfo.ColumnFields
        .Cells(lastRow + 1, 9) = ptInfo.PageFields
        .Cells(lastRow + 1, 10) = ptInfo.TotalFields
        .Cells(lastRow + 1, 11) = ptInfo.CalculatedItems
        .Cells(lastRow + 1, 12) = ptInfo.CalculatedFields
    End With
End Sub
Sub OutputDataFieldInfo(ByRef dfInfo As DataFieldInfo)
    Dim wks As Worksheet, lastRow As Long
    Set wks = ThisWorkbook.Sheets("DataFieldInfo")
    lastRow = LastRowInOneColumn("A", wks)
    With wks
        .Cells(lastRow + 1, 1) = dfInfo.Workbook
        .Cells(lastRow + 1, 2) = dfInfo.Worksheet
        .Cells(lastRow + 1, 3) = dfInfo.PivotTable
        .Cells(lastRow + 1, 4) = dfInfo.Name
        .Cells(lastRow + 1, 5) = dfInfo.Aggregate
    End With
End Sub
