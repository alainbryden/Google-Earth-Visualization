Attribute VB_Name = "DataRangesRoutines"
Option Explicit

'The names of named ranges in this workbook
Public Const rngColours = "Colours" 'Colours that can be used for the bar graph are listed here
Public Const rngColumnHeaders = "ColumnHeaders" 'The names of columns in the exposure information sheet
Public Const rngColumnHeadersPlus = "ColumnHeadersPlus" 'Same as above, but with an added field '(automatic)'
Public Const rngAvailableCategories = "AvailableCategories" 'List of categories selected in config to be available to the user
Public Const rngAvailablePlotValues = "AvailablePlotValues" 'List of plot values selected to be available to the user
Public Const rngPlotterCategories = "PlotterCategories" 'List of unique categories found in the selected plot category column
Public Const rngLocatorCategories = "LocatorCategories" 'List of unique categories found in the "View all Locations" selected category column
Public Const rngFilterCopyRange = "FilterCopyRange" 'The area where a filter's results are temporarily copied

'Abstracted routine that will perform any updates that might be necessary
'to ensure the ranges contain up to date data..
Public Sub UpdateRanges()
    RefreshColumnHeaders
    DataRanges.Columns.AutoFit
End Sub

'Updates the list of columns that can be found in the exposure information sheet
Private Sub RefreshColumnHeaders()
    Dim col As Integer
    With DataRanges
        col = .Range(rngColumnHeaders).column 'Get the column number containing the named range
        .Columns(col).ClearContents 'Clear the current list of column headers
        toEndOfRow(wkstLocations.Cells(1, 1)).Copy 'Copy the column names from the exposure information sheet
        'Paste (transpose) the column names into the DataRanges sheet
        .Cells(1, col).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=True
        Application.CutCopyMode = False
        'Set the named ranges
        Names(rngColumnHeaders).RefersTo = toEndOfColumn(.Cells(1, col))
        'Create one more header '(automatic)' which is can be selected in some cases to auto-generate a new column.
        endOfColumn(DataRanges, col).Offset(1, 0) = "(Automatic)"
        'Set the named ranges that has the added header
        Names(rngColumnHeadersPlus).RefersTo = toEndOfColumn(.Cells(1, col))
    End With
End Sub

'Get the columns selected in advanced configuration into a named range which is bound
'to listboxes meant to contain all available categorizations.
Public Sub RefreshSelectedCategorizations()
    Dim arrItems() As String, i As Integer
    With lbCategorizations()
        If .listCount = 0 Then
            ReDim arrItems(0 To 0)
        Else
            ReDim arrItems(LBound(.List) - 1 To UBound(.List))
        End If
        arrItems(LBound(arrItems)) = "(No Categorization)"
        For i = LBound(arrItems) + 1 To UBound(arrItems)
            arrItems(i) = .List(i)
        Next i
    End With
    DefineNamedRangeContents arrItems, rngAvailableCategories
End Sub

'Get the columns selected in advanced configuration into a named range which is bound
'to listboxes meant to contain all available plot values.
Public Sub RefreshSelectedPlotValues()
    DefineNamedRangeContents lbPlotValues().List, rngAvailablePlotValues
End Sub

'Fills the column associated with a named range with the specified list of values and fixes the named range
Private Sub DefineNamedRangeContents(listItems As Variant, destRangeName As String)
    Dim i As Integer, itemCount As Integer, col As Integer
    If IsNull(listItems) Then itemCount = 0 Else itemCount = UBound(listItems) - LBound(listItems) + 1
    With DataRanges
        col = .Range(destRangeName).column 'Get the column number containing the named range
        .Columns(col).ClearContents 'Clear the contents of the column containing the named range
        'Update the named range to refer to the new column names
        If itemCount = 0 Then
            Names(destRangeName).RefersTo = .Cells(1, col)
        Else
            .Range(.Cells(1, col), .Cells(itemCount, col)).value = Application.Transpose(listItems)
            Names(destRangeName).RefersTo = .Range(.Cells(1, col), .Cells(itemCount, col))
        End If
    End With
End Sub

'Makes the list of categories to appear in the dropdown box for viewing multiple locations by category
Public Sub RefreshViewLocationsCategories(ByVal ColumnName As String)
    Dim columnValues() As String, valuesForRange() As String
    columnValues = GetUniqueCategories(ColumnName)
    If (Not Not columnValues) = 0 Then
        ReDim valuesForRange(0 To 0)
    Else
        ReDim valuesForRange(LBound(columnValues) - 1 To UBound(columnValues))
    End If
    valuesForRange(LBound(valuesForRange)) = "(All Remaining Locations)"
    If (Not Not columnValues) <> 0 Then
        Dim i As Integer
        For i = LBound(columnValues) To UBound(columnValues)
            valuesForRange(i) = IIf(Trim(columnValues(i)) = "", "(Blank)", columnValues(i))
        Next i
    End If
    DefineNamedRangeContents valuesForRange, rngLocatorCategories
End Sub

'Makes the list of categories to appear in the dropdown boxes defining a per-category colour scheme for a bar graph
Public Sub RefreshPlotterCategories(ByVal ColumnName As String)
    Dim columnValues() As String, valuesForRange() As String
    columnValues = GetUniqueCategories(ColumnName)
    If (Not Not columnValues) = 0 Then  'Returns 0 only if uninitialized
        ReDim valuesForRange(0 To 0)
    Else
        ReDim valuesForRange(LBound(columnValues) - 1 To UBound(columnValues))
    End If
    valuesForRange(LBound(valuesForRange)) = "(All Remaining Locations)"
    If (Not Not columnValues) <> 0 Then
        Dim i As Integer
        For i = LBound(columnValues) To UBound(columnValues)
            valuesForRange(i) = IIf(Trim(columnValues(i)) = "", "(Blank)", columnValues(i))
        Next i
    End If
    DefineNamedRangeContents valuesForRange, rngPlotterCategories
End Sub

'Get the set of unique categories which exist in the column selected as the current categorization.
Private Function GetUniqueCategories(ByVal ColumnName As String) As String()
    Dim srcRange As Range
    Dim dstCol As Integer, srcCol As Integer, i As Integer
    Dim Result() As String
    dstCol = DataRanges.Range(rngFilterCopyRange).column
    srcCol = getColumnNumber(ColumnName)
    If srcCol <> 0 Then
        With DataRanges
            .Columns(dstCol).ClearContents
            'Filter the selected column and find all unique categories within
            Set srcRange = toEndOfColumn(wkstLocations.Cells(1, srcCol))
            srcRange.AdvancedFilter xlFilterCopy, , .Cells(1, dstCol), True
            .Cells(1, dstCol).Delete xlUp
            ReDim Result(1 To toEndOfColumn(.Cells(1, dstCol)).Count)
            toEndOfColumn(.Cells(1, dstCol)).SortSpecial Header:=xlNo
            For i = LBound(Result) To UBound(Result)
                Result(i) = .Cells(i, dstCol).Text
            Next i
            .Columns(dstCol).ClearContents
        End With
    End If
    GetUniqueCategories = Result
End Function
