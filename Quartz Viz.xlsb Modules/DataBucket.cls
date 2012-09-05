VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataBucket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TotalValue As Double
Private ValueCount As Double
Public i As Integer, j As Integer, k As Integer

Public Property Get Average() As Double
    If ValueCount > 0 Then
        Average = TotalValue / ValueCount
    Else
        Average = 0
    End If
End Property

Public Property Get Sum() As Double
    Sum = TotalValue
End Property

Public Property Get Count() As Double
    Count = ValueCount
End Property

Public Function Add(value As Double)
    TotalValue = TotalValue + value
    ValueCount = ValueCount + 1
End Function
