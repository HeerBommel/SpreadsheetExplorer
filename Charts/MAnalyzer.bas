Attribute VB_Name = "MAnalyzer"
Option Explicit

Sub AnalyzeWorkbook(ByVal FileName As String)
    Dim wkb As Workbook, errorInfo As ErrorLog, cInfo As ChartInfo
    On Error Resume Next
    Set wkb = Application.Workbooks.Open(FileName:=PATH & FileName, UpdateLinks:=False, ReadOnly:=True)
    On Error GoTo 0
    If Not wkb Is Nothing Then
        Application.StatusBar = FileName
        If wkb.Charts.Count > 0 Then 'if there are any chart sheets
            Dim c As Chart
            For Each c In wkb.Charts
                Set cInfo = GetChartInfo(c)
                cInfo.FileName = FileName
                OutputChartInfo cInfo
            Next c
        End If
        Dim wks As Worksheet
        For Each wks In wkb.Worksheets
            If wks.ChartObjects.Count > 0 Then 'if there are any embedded charts
                Dim co As ChartObject
                For Each co In wks.ChartObjects
                    Set cInfo = GetChartInfo(co.Chart)
                    cInfo.FileName = FileName
                    OutputChartInfo cInfo
                Next co
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
Function GetChartInfo(ByRef c As Chart) As ChartInfo
    Dim cInfo As ChartInfo
    Set cInfo = New ChartInfo
    If TypeName(c.Parent) = "ChartObject" Then
        cInfo.Emmbedded = "Yes"
        cInfo.Index = c.Parent.Index
        cInfo.Worksheet = c.Parent.Parent.Name
    Else
        cInfo.Emmbedded = "No"
        cInfo.Index = c.Index
        cInfo.Worksheet = c.Name
    End If
    cInfo.Name = c.Name
    If c.HasTitle Then
        cInfo.Title = c.ChartTitle.Text
    End If
    If chartTypes.Exists(c.ChartType) Then
        cInfo.ChartType = chartTypes.Item(c.ChartType)
    Else
        cInfo.ChartType = c.ChartType
    End If
    Set GetChartInfo = cInfo
End Function
