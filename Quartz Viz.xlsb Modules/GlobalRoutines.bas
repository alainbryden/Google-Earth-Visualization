Attribute VB_Name = "GlobalRoutines"
Option Explicit

'Has to be run manually sometimes when stopping excel in the middle of a routine.
Public Function fixScreenUpdating()
    Automation False
End Function

'Rounds a number to the correct number of significant figures
Public Function sigFigs(num As Variant, sigs As Integer)
    Dim exponent As Double
    If IsNumeric(num) And sigs > 0 Then
        exponent = IIf(num <> 0, Int(Log(Abs(num)) / Log(10#)), 0)
        sigFigs = WorksheetFunction.Round(num, sigs - (1 + exponent))
    Else: sigFigs = num
    End If
End Function

'For large automation tasks, configures the workbook for speedup and clean appearance
Public Sub Automation(ByVal automationActive As Boolean)
    Application.ScreenUpdating = Not automationActive
    Application.EnableEvents = Not automationActive
    Application.Calculation = IIf(automationActive, xlCalculationManual, xlCalculationAutomatic)
    Application.StatusBar = False
End Sub

'Looks up the column number based on the name of the column header
Public Function getColumnNumber(ByVal ColumnName As String) As Integer
    On Error GoTo noMatch
    getColumnNumber = Application.WorksheetFunction.Match(ColumnName, wkstLocations.Rows(1), 0)
    Exit Function
noMatch:
    getColumnNumber = 0
End Function

'Routine for improved concatenation of strings
Public Sub fastConcat(ByRef Dest As String, ByRef Source As String, _
                      ByRef ccOffset As Long, Optional ByVal ccDoubleRate As Double = 2)
    Dim ccIncrement As Long
    If ccDoubleRate < 1 Then ccDoubleRate = 1
    ccIncrement = ccDoubleRate * Len(Dest)
    Dim l As Long: l = Len(Source)
    If (ccOffset + l) >= Len(Dest) Then
        Dest = Dest & Space$(IIf(l > ccIncrement, l, ccIncrement))
    End If
    Mid$(Dest, ccOffset + 1, l) = Source
    ccOffset = ccOffset + l
End Sub

'Get the range from the selected cell to the last cell in a column
Public Function toEndOfColumn(ByRef firstCell As Range) As Range
    With firstCell
        Set toEndOfColumn = .Worksheet.Range(firstCell, endOfColumn(.Worksheet, .column))
    End With
End Function

'Get the range from the selected cell to the last cell in a row
Public Function toEndOfRow(ByRef firstCell As Range) As Range
    With firstCell
        Set toEndOfRow = .Worksheet.Range(firstCell, endOfRow(.Worksheet, .row))
    End With
End Function

'Get the last cell in the specified column
Public Function endOfColumn(ByRef srcWorkSheet As Worksheet, ByVal columnNumber As Integer) As Range
    Set endOfColumn = srcWorkSheet.Cells(srcWorkSheet.Rows.Count, columnNumber).End(xlUp)
End Function

'Get the last cell in the specified row
Public Function endOfRow(ByRef srcWorkSheet As Worksheet, ByVal rowNumber As Integer) As Range
    Set endOfRow = srcWorkSheet.Cells(rowNumber, srcWorkSheet.Columns.Count).End(xlToLeft)
End Function

'Removes characters that shouldn't exist in the kml file.
Public Function xmlFormatString(ByVal strIn As String) As String
    Dim original As String
    original = strIn
    original = Replace(original, "&", "&amp;")
    original = Replace(original, "<", "&lt;")
    original = Replace(original, ">", "&gt;")
    xmlFormatString = ""
    Dim c$, i As Integer
    For i = 1 To Len(original)
        c$ = Mid(original, i, 1)
        xmlFormatString = xmlFormatString & IIf(c$ Like "[0-9A-Za-z &;]", c$, "&#" & Asc(c$) & ";")
    Next i
End Function

'Gets the interior colour of a cell regardless of conditional formatting
Function ConditionalColor(rg As Range) As Long
    Dim cel As Range
    Dim tmp As Variant
    Dim boo As Boolean
    Dim frmla As String, frmlaR1C1 As String, frmlaA1 As String
    Dim i As Long
     
    Set cel = rg.Cells(1, 1)
    ConditionalColor = cel.Interior.color
     
    If cel.FormatConditions.Count > 0 Then
        With cel.FormatConditions
            For i = 1 To .Count 'Loop through the possible format conditions for each cell
                frmla = .Item(i).Formula1
                If Left(frmla, 1) = "=" Then
                    frmlaR1C1 = Application.ConvertFormula(frmla, xlA1, xlR1C1, , cel)
                    frmlaA1 = Application.ConvertFormula(frmlaR1C1, xlR1C1, xlA1, xlAbsolute, cel)
                    boo = Application.Evaluate(frmlaA1)
                Else
                    Select Case .Item(i).Operator
                        Case xlEqual: frmla = cel & "=" & .Item(i).Formula1
                        Case xlNotEqual: frmla = cel & "<>" & .Item(i).Formula1
                        Case xlBetween: frmla = "AND(" & .Item(i).Formula1 & "<=" & cel & "," & cel & "<=" & .Item(i).Formula2 & ")"
                        Case xlNotBetween: frmla = "OR(" & .Item(i).Formula1 & ">" & cel & "," & cel & ">" & .Item(i).Formula2 & ")"
                        Case xlLess: frmla = cel & "<" & .Item(i).Formula1
                        Case xlLessEqual: frmla = cel & "<=" & .Item(i).Formula1
                        Case xlGreater: frmla = cel & ">" & .Item(i).Formula1
                        Case xlGreaterEqual: frmla = cel & ">=" & .Item(i).Formula1
                    End Select
                    boo = Application.Evaluate(frmla)
                End If
                 
                If boo Then 'If this Format Condition is satisfied
                    On Error Resume Next
                    tmp = .Item(i).Interior.color
                    If Err = 0 Then ConditionalColor = tmp
                    Err.Clear
                    On Error GoTo 0
                    Exit For 'Since Format Condition is satisfied, exit the inner loop
                End If
            Next i
        End With
    End If
End Function
