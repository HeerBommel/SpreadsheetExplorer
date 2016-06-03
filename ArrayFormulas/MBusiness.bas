Attribute VB_Name = "MBusiness"
Option Compare Text
Option Explicit
Const FOLDER_TO_ANALYZE As String = "X:\TUDelft\E\Enron\converted\To Scan\"
Const FOLDER_TO_SAVE As String = "X:\TUDelft\S\SEMS16\ConainsArrayFormulas\"
Dim totalWkb As Long

Sub Demo()
    Dim file As String
    
    Application.Calculation = xlCalculationManual
    file = Dir(FOLDER_TO_ANALYZE)
    Do While file <> ""
        AnalyzeFile FOLDER_TO_ANALYZE & file
        file = Dir
    Loop
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub

Sub AnalyzeFile(ByVal fullPath As String)
    Dim wkb As Workbook, c As Variant, wks As Worksheet, noChartSheets As Long, noArrayFormulas As Long
    Dim category As String, fileName As String, wksData As Worksheet, lastRow As Long
    Dim errorDescription As String
    Application.DisplayAlerts = False: Application.ScreenUpdating = False: Application.EnableEvents = False
    Set wksData = ThisWorkbook.Sheets("SheetInfo")
    totalWkb = totalWkb + 1
    Application.StatusBar = totalWkb
    On Error Resume Next
    Set wkb = Application.Workbooks.Open(fileName:=fullPath, ReadOnly:=True, UpdateLinks:=False)
    errorDescription = Err.Description
    lastRow = LastRowInOneColumn("A", wksData)
    On Error GoTo 0
    wksData.Cells(lastRow + 1, 1) = Now
    If wkb Is Nothing Then
        wksData.Cells(lastRow + 1, 2) = fullPath
        wksData.Cells(lastRow + 1, 5) = errorDescription
        wksData.Cells(lastRow + 1, 4) = "ERROR"
    End If
    If Not wkb Is Nothing Then
        category = Mid(fullPath, 20, InStr(20, fullPath, "\") - 20)
        fileName = Split(fullPath, "\")(UBound(Split(fullPath, "\")))
        'fileName = Left(fileName, InStrRev(fileName, ".") - 1)
        Dim formulas As Range, f As Range, dctArray As Dictionary, k As Variant
        For Each wks In wkb.Worksheets
            On Error Resume Next
            Set formulas = wks.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            Set dctArray = New Dictionary
            If Not formulas Is Nothing Then
                For Each f In formulas
                    If f.HasArray Then
                        If Not dctArray.Exists(f.FormulaR1C1) Then
                            dctArray.Add f.FormulaR1C1, f.FormulaArray
                            OutputArrayFormulas GetFormulaInfo(f)
                        End If
                    End If
                Next f
                For Each k In dctArray.Keys
                    Debug.Print k
                Next k
            End If
            noArrayFormulas = noArrayFormulas + dctArray.Count
        Next wks
        If noArrayFormulas > 0 Then
            FileCopy fullPath, FOLDER_TO_SAVE & category & "_" & Split(fullPath, "\")(UBound(Split(fullPath, "\")))
            wksData.Cells(lastRow + 1, 4) = "YES"
        Else
            wksData.Cells(lastRow + 1, 4) = "NO"
        End If
        wksData.Cells(lastRow + 1, 2) = FOLDER_TO_SAVE & category & "_" & Split(fullPath, "\")(UBound(Split(fullPath, "\")))
        wksData.Cells(lastRow + 1, 3) = noArrayFormulas
        wkb.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = True: Application.ScreenUpdating = True: Application.EnableEvents = False
End Sub
Function GetFormulaInfo(ByRef f As Range) As ArrayFormula
    Dim af As ArrayFormula
    Set af = New ArrayFormula
    af.fileName = f.Parent.Parent.Name
    af.sheetName = f.Parent.Name
    af.cellAddress = f.Address
    af.formula = f.formula
    Set GetFormulaInfo = af
End Function
