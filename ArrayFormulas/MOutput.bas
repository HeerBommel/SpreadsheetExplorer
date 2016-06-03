Attribute VB_Name = "MOutput"
Option Explicit

Sub OutputArrayFormulas(ByRef af As ArrayFormula)
    Dim wks As Worksheet, lastRow As Long
    Set wks = ThisWorkbook.Sheets("FormulaInfo")
    lastRow = LastRowInOneColumn("A", wks)
    wks.Cells(lastRow + 1, 1) = af.fileName
    wks.Cells(lastRow + 1, 2) = af.sheetName
    wks.Cells(lastRow + 1, 3) = af.cellAddress
    wks.Cells(lastRow + 1, 4) = "'" & af.formula
End Sub
