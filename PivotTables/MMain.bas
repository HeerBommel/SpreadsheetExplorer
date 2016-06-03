Attribute VB_Name = "MMain"
Option Explicit
Sub Main()
    PerformanceMode True
    initGlobals
    ClearOutputSheet
    currentFile = Dir(PATH & EXT)
    Do While currentFile <> ""
        AnalyzeWorkbook currentFile
        currentFile = Dir
    Loop
    PerformanceMode False
    Application.StatusBar = False
End Sub

