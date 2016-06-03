VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Workbook As String
Public Worksheet As String
Public Name As String
Public memory As Long
Public records As Long
Public DataFields As Integer
Public RowFields As Integer
Public ColumnFields As Integer
Public PageFields As Integer
Public TotalFields As Integer
Public CalculatedItems As Integer
Public CalculatedFields As Integer


