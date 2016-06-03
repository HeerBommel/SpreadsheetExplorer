Attribute VB_Name = "MAnalyzer"
Option Explicit

Sub AnalyzeWorkbook(ByVal FileName As String)
    Dim wkb As Workbook, errorInfo As ErrorLog, ptInfo As PivotInfo, df As PivotField, dfInfo As DataFieldInfo
    On Error Resume Next
    Set wkb = Application.Workbooks.Open(FileName:=PATH & FileName, UpdateLinks:=False, ReadOnly:=True)
    On Error GoTo 0
    If Not wkb Is Nothing Then
        Application.StatusBar = FileName
        Dim wks As Worksheet
        For Each wks In wkb.Worksheets
            If wks.PivotTables.Count > 0 Then 'if there are any pivot tables
                Dim pt As PivotTable
                For Each pt In wks.PivotTables
                    Set ptInfo = GetPivotInfo(pt)
                    OutputPivotInfo ptInfo
                    For Each df In pt.DataFields
                        Set dfInfo = GetDataFieldInfo(df, pt)
                        OutputDataFieldInfo dfInfo
                    Next df
                Next pt
            End If
        Next wks
        wkb.Close SaveChanges:=False
    Else
        Set errorInfo = New ErrorLog
        errorInfo.ErrorCode = "Could not open file"
        errorInfo.info = currentFile
        OutputErrors errorInfo
    End If
End Sub
Function GetPivotInfo(ByRef pt As PivotTable) As PivotInfo
    Dim ptInfo As PivotInfo
    Set ptInfo = New PivotInfo
    With ptInfo
        .Workbook = pt.Parent.Parent.Name
        .Worksheet = pt.Parent.Name
        .Name = pt.Name
        .memory = GetPivotMemory(pt)
        .records = GetPivotRecords(pt)
        .DataFields = pt.DataFields.Count
        .RowFields = pt.RowFields.Count
        .ColumnFields = pt.ColumnFields.Count
        .PageFields = pt.PageFields.Count
        .TotalFields = pt.PivotFields.Count
        .CalculatedItems = GetCalculatedItems(pt)
        .CalculatedFields = pt.CalculatedFields.Count
    End With
    Set GetPivotInfo = ptInfo
End Function
Function GetDataFieldInfo(ByRef df As PivotField, ByRef pt As PivotTable) As DataFieldInfo
    Dim dfInfo As DataFieldInfo
    Set dfInfo = New DataFieldInfo
    With dfInfo
        .Workbook = pt.Parent.Parent.Name
        .Worksheet = pt.Parent.Name
        .PivotTable = pt.Name
        .Name = df.Name
        .Aggregate = df.Function
    End With
    Set GetDataFieldInfo = dfInfo
End Function
Function GetPivotMemory(ByRef pt As PivotTable) As Long
    Dim memory As Long
    On Error Resume Next
    memory = pt.Parent.Parent.PivotCaches(pt.CacheIndex).memoryUsed
    On Error GoTo 0
    If Err.Number = 0 Then
        GetPivotMemory = memory
    Else
        GetPivotMemory = 0
    End If
End Function
Function GetPivotRecords(ByRef pt As PivotTable) As Long
    Dim records As Long
    On Error Resume Next
    records = pt.Parent.Parent.PivotCaches(pt.CacheIndex).RecordCount
    On Error GoTo 0
    If Err.Number = 0 Then
        GetPivotRecords = records
    Else
        GetPivotRecords = 0
    End If
End Function

Function GetCalculatedItems(ByRef pt As PivotTable) As Integer
    Dim f As PivotField, noCI As Integer, itemsInField As Integer
    noCI = 0
    For Each f In pt.PivotFields
        On Error Resume Next
        itemsInField = f.CalculatedItems.Count
        On Error GoTo 0
        If itemsInField > 0 Then noCI = noCI + 1
    Next f
    GetCalculatedItems = noCI
End Function
