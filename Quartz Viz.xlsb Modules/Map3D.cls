VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Map3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'A 3D construction where each member of the main collection (X)
'is a collection (Y), which contains the final objects at various (Z)
'for that X, Y key.
Private main As New Dictionary
Private elements As Long 'Number of elements in this collection
Private iterator_count As Long 'Used when iterating over elements in the collection

'Adds an object with the given x, y, z coordinates as a key
Public Sub Add(ByRef obj As Variant, ByVal x As Long, ByVal y As Long, ByVal z As Long)
    Dim Ycol As Dictionary, Zcol As Dictionary

    If main.Exists(x) Then
        Set Ycol = main(x)
    Else
        Set Ycol = New Dictionary
        main.Add x, Ycol
    End If
    
    If Ycol.Exists(y) Then
        Set Zcol = Ycol(y)
    Else
        Set Zcol = New Dictionary
        Ycol.Add y, Zcol
    End If
    
    Zcol.Add z, obj
    elements = elements + 1
End Sub

Public Function Exists(ByVal x As Long, ByVal y As Long, ByVal z As Long)
    If main.Exists(x) Then If main(x).Exists(y) Then If main(x)(y).Exists(z) Then Exists = True
End Function

'Gets an array of items in the data structure from lowest to highest x, y, then z indexes
Public Function toArray() As Variant
    Dim arrayElems() As Variant
    Dim Ycol As Dictionary, Zcol As Dictionary
    Dim i As Long, j As Long, k As Long, elem As Long
    Dim iKeys(), jKeys(), kKeys()
    
    ReDim arrayElems(0 To elements - 1)
    elem = 0
    
    iKeys = main.Keys()
    Sort iKeys
    For i = LBound(iKeys) To UBound(iKeys)
        Set Ycol = main(iKeys(i))
        jKeys = Ycol.Keys()
        Sort jKeys
        For j = LBound(jKeys) To UBound(jKeys)
            Set Zcol = Ycol(jKeys(j))
            kKeys = Zcol.Keys()
            Sort kKeys
            For k = LBound(kKeys) To UBound(kKeys)
                Set arrayElems(elem) = Zcol(kKeys(k))
                elem = elem + 1
            Next k
        Next j
    Next i
    
    toArray = arrayElems
End Function

'Retrieves an object with the x, y, z coordinate key specified
Public Function Item(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Variant
    On Error Resume Next
    Set Item = main(x)(y)(z)
    'If this error occured (type mismatch) then the element is a non-object type so don't use the set keyword
    If Err.Number = 13 Then On Error GoTo 0: Item = main(x)(y)(z)
End Function

'Retrieves the collection of objects stored at the x, y coordinates specified
Public Function Items(ByVal x As Long, ByVal y As Long) As Dictionary
    Set Items = main(x)(y)
End Function

'Deletes an object with the x, y, z coordinate key specified
Public Sub Remove(ByVal x As Long, ByVal y As Long, ByVal z As Long)
    Dim Ycol As Dictionary, Zcol As Dictionary
    Set Ycol = main(x)
    Set Zcol = Ycol(y)
    Zcol.Remove z
    If Zcol.Count = 0 Then Ycol.Remove y
    If Ycol.Count = 0 Then main.Remove x
    elements = elements - 1
End Sub

'Get the number of elements in the list
Public Property Get Count() As Long
    Count = elements
End Property
