Attribute VB_Name = "QuickSort"
Option Explicit
Private Sub QuickSort(ByRef a(), ByVal l As Long, ByVal r As Long)
    Dim M As Long, i As Long, j As Long, v As Long
    M = 4

    If ((r - l) > M) Then
        i = (r + l) / 2
        If (a(l) > a(i)) Then swap a, l, i
        If (a(l) > a(r)) Then swap a, l, r
        If (a(i) > a(r)) Then swap a, i, r

        j = r - 1
        swap a, i, j
        i = l
        v = a(j)
        Do
            Do: i = i + 1: Loop While (a(i) < v)
            Do: j = j - 1: Loop While (a(j) > v)
            If (j < i) Then Exit Do
            swap a, i, j
        Loop
        swap a, i, r - 1
        QuickSort a, l, j
        QuickSort a, i + 1, r
    End If
End Sub

Private Sub swap(ByRef a(), ByVal i As Long, ByVal j As Long)
    Dim T
    T = a(i)
    a(i) = a(j)
    a(j) = T
End Sub

Private Sub InsertionSort(ByRef a(), ByVal lo0 As Long, ByVal hi0 As Long)
    Dim i As Long, j As Long, v As Long

    For i = lo0 + 1 To hi0
        v = a(i)
        j = i
        Do While j > lo0
            If Not a(j - 1) > v Then Exit Do
            a(j) = a(j - 1)
            j = j - 1
        Loop
        a(j) = v
    Next i
End Sub

Public Sub Sort(ByRef a())
    QuickSort a, LBound(a), UBound(a)
    InsertionSort a, LBound(a), UBound(a)
End Sub

