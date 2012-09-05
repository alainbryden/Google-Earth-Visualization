Attribute VB_Name = "ConfigControlsRoutines"
Option Explicit

'The names given to various controls on the config sheet
Public Const ctrl_lbCategorizations = "lbCategorizations"
Public Const ctrl_ddCategorizations = "ddCategorizations"
Public Const ctrl_lbPlotValues = "lbPlotValues"
Public Const ctrl_ddPlotValues = "ddPlotValues"
Public Const ctrl_ddUniqueIdentifier = "ddUniqueIdentifier"
Public Const ctrl_ddLattitude = "ddLattitude"
Public Const ctrl_ddLongitude = "ddLongitude"
Public Const ctrl_ddLookUpAddress = "ddLookUpAddress"
Public Const ctrl_ddDescription = "ddDescription"
'Names given to named ranges containing the column number associated with various controls
Public Const rng_Longitude = "rngColLongitude"
Public Const rng_Lattitude = "rngColLattitude"
Public Const rng_LookUpAddress = "rngColLookupAddress"
Public Const rng_Description = "rngColDescription"

'Functions for returning the control objects in the spreadsheet.
Public Function lbCategorizations() As ListBox
    Set lbCategorizations = Config.ListBoxes(ctrl_lbCategorizations)
End Function
Public Function ddCategorizations() As DropDown
    Set ddCategorizations = Config.DropDowns(ctrl_ddCategorizations)
End Function
Public Function lbPlotValues() As ListBox
    Set lbPlotValues = Config.ListBoxes(ctrl_lbPlotValues)
End Function
Public Function ddPlotValues() As DropDown
    Set ddPlotValues = Config.DropDowns(ctrl_ddPlotValues)
End Function
Public Function ddUniqueIdentifier() As DropDown
    Set ddUniqueIdentifier = Config.DropDowns(ctrl_ddUniqueIdentifier)
End Function
Public Function ddLattitude() As DropDown
    Set ddLattitude = Config.DropDowns(ctrl_ddLattitude)
End Function
Public Function ddLongitude() As DropDown
    Set ddLongitude = Config.DropDowns(ctrl_ddLongitude)
End Function
Public Function ddLookUpAddress() As DropDown
    Set ddLookUpAddress = Config.DropDowns(ctrl_ddLookUpAddress)
End Function
Public Function ddDescription() As DropDown
    Set ddDescription = Config.DropDowns(ctrl_ddDescription)
End Function

'Changes the unique identifier used to refer to rows from the exposure sheet
Public Sub ddUniqueIdentifier_Change()
    Automation True
    ToolsControlsRoutines.UpdateControls
    Automation False
End Sub

Private Sub AddCategorization_Click()
    AddColumnParamed lbCategorizations(), ddCategorizations()
    DataRangesRoutines.RefreshSelectedCategorizations
End Sub

Private Sub RemoveCategorization_Click()
    RemoveColumnParamed lbCategorizations()
    DataRangesRoutines.RefreshSelectedCategorizations
End Sub

Private Sub CategorizationIndexUp_Click()
    ColumnIndexUpParamed lbCategorizations()
    DataRangesRoutines.RefreshSelectedCategorizations
End Sub

Private Sub CategorizationIndexDown_Click()
    ColumnIndexDownParamed lbCategorizations()
    DataRangesRoutines.RefreshSelectedCategorizations
End Sub

Private Sub AddPlotValue_Click()
    AddColumnParamed lbPlotValues(), ddPlotValues()
    DataRangesRoutines.RefreshSelectedPlotValues
End Sub

Private Sub RemovePlotValue_Click()
    RemoveColumnParamed lbPlotValues()
    DataRangesRoutines.RefreshSelectedPlotValues
End Sub

Private Sub PlotValueIndexUp_Click()
    ColumnIndexUpParamed lbPlotValues()
    DataRangesRoutines.RefreshSelectedPlotValues
End Sub

Private Sub PlotValueIndexDown_Click()
    ColumnIndexDownParamed lbPlotValues()
    DataRangesRoutines.RefreshSelectedPlotValues
End Sub

'This routine is called when the user clicks the '+' button to add a
'column to one of the lists of columns in the configuration pages.
'First, we check if the column is already in the list. If not, we add it.
Private Sub AddColumnParamed(ByRef listBoxColumns As ListBox, ByRef dropDownColumns As DropDown)
    Dim valueExists As Boolean: valueExists = False
    Dim i As Variant, valueToAdd As String
    valueToAdd = dropDownColumns.List(dropDownColumns.value)
    If Not listBoxColumns.listCount = 0 Then
        For Each i In listBoxColumns.List
            If i = valueToAdd Then valueExists = True: Exit For
        Next i
    End If
    If Not valueExists Then listBoxColumns.AddItem (valueToAdd)
End Sub

'This removes the column name that is currently selected in a listbox of added
'columns. This happens when the user clicks the 'X' button.
Private Sub RemoveColumnParamed(ByRef listBoxColumns As ListBox)
    If listBoxColumns.ListIndex <> 0 Then _
        listBoxColumns.RemoveItem (listBoxColumns.ListIndex)
End Sub

'Changes the order of items in a column list. Excel doesn't allow manipulating
'list items directly, so we must remove all items and add them back in the correct order.
Private Sub ColumnIndexUpParamed(ByRef listBoxColumns As ListBox)
    Dim i As Integer
    Dim temp As Variant
    With listBoxColumns
        i = .ListIndex
        If i > 1 Then
            .Selected(i) = False
            temp = .List(i - 1)
            .List(i - 1) = .List(i)
            .List(i) = temp
            .Selected(i - 1) = True
        End If
    End With
End Sub

'Changes the order of items in a column list. Excel doesn't allow manipulating
'list items directly, so we must remove all items and add them back in the correct order.
Private Sub ColumnIndexDownParamed(ByRef listBoxColumns As ListBox)
    Dim i As Integer
    Dim temp As Variant
    With listBoxColumns
        i = .ListIndex
        If i < .listCount Then
            .Selected(i) = False
            temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = temp
            .Selected(i + 1) = True
        End If
    End With
End Sub

'If for some reason these ranges aren't updated, this forces the controls to be rebuilt based on the current
'state of the exposure information sheet.
Sub ScanExposureInformationSheet()
    wkstLocations.Activate
    wkstLocations.somethingChanged = True
    Config.Activate
End Sub


'Make sure the controls on the config sheet contain only valid entries
Public Sub UpdateControls()
    validateColumnReferences lbCategorizations()
    validateColumnReferences lbPlotValues()
    'Ensure the underlying column numbers match the selected values in the drop down boxes
    ddLattitude_Change
    ddLongitude_Change
    ddLookupAddress_Change
    ddDescription_Change
    'If any categories were deleted, we want these changes reflected in the associated ranges
    DataRangesRoutines.RefreshSelectedCategorizations
    DataRangesRoutines.RefreshSelectedPlotValues
End Sub

'Can remove column references from a list when the columns no longer exist
'(for instance, if the tool is configured for a differently formatted exposure set)
'Will prompt the user about invalid column references and display the current configuration
Public Sub validateColumnReferences(ByRef listBoxColumns As ListBox)
    Dim userPrompted As Boolean: userPrompted = False
    Dim currIndex As Variant
    
    If Not listBoxColumns.listCount = 0 Then
        Dim i As Integer: i = 1
        For Each currIndex In listBoxColumns.List
            If getColumnNumber(currIndex) = 0 Then
                If Not userPrompted Then
                    If MsgBox("Some selected columns cannot be found in " & _
                              "the current exposure set. Do you wish to remove invalid entries?", _
                              vbYesNo, "Invalid columns") = vbNo Then GoTo skipRemovals
                    userPrompted = True
                End If
                listBoxColumns.RemoveItem (i): i = i - 1
            End If
            i = i + 1
        Next currIndex
    End If
    Exit Sub
skipRemovals:
    Config.Visible = True
    Config.Activate 'Display the current configuration
End Sub

'Get the unique column number
Public Function colUniqueID() As Integer
    colUniqueID = getColumnFor(ctrl_ddUniqueIdentifier, False)
End Function

'Get the lattitude column number
Public Function colLattitude(Optional ByVal createIfMissing = False) As Integer
    colLattitude = getColumnFor(ctrl_ddLattitude, createIfMissing)
End Function

'Get the longitude column number
Public Function colLongitude(Optional ByVal createIfMissing = False) As Integer
    colLongitude = getColumnFor(ctrl_ddLongitude, createIfMissing)
End Function

'Get the LookupAddress column number
Public Function colLookupAddress() As Integer
    colLookupAddress = getColumnFor(ctrl_ddLookUpAddress, False)
End Function

'Get the Description column number
Public Function colDescription() As Integer
    colDescription = getColumnFor(ctrl_ddDescription, False)
End Function

Public Sub ddLattitude_Change()
    Dim prevProtection: prevProtection = Config.ProtectionMode: Config.Unprotect
    Config.Range(rng_Lattitude).value = colLattitude()
    If prevProtection = True Then Config.Protect
End Sub

Public Sub ddLongitude_Change()
    Dim prevProtection: prevProtection = Config.ProtectionMode: Config.Unprotect
    Config.Range(rng_Longitude).value = colLongitude()
    If prevProtection = True Then Config.Protect
End Sub

Public Sub ddLookupAddress_Change()
    Dim prevProtection: prevProtection = Config.ProtectionMode: Config.Unprotect
    Config.Range(rng_LookUpAddress).value = colLookupAddress()
    If prevProtection = True Then Config.Protect
End Sub

Public Sub ddDescription_Change()
    Dim prevProtection: prevProtection = Config.ProtectionMode: Config.Unprotect
    Config.Range(rng_Description).value = colDescription()
    If prevProtection = True Then Config.Protect
End Sub

'Gets the column number of a key column type. If none is assigned, it creates one and then returns
'the number of the new column. Also adjusts config to point to the correct column.
Private Function getColumnFor(ByVal dropDownBoxName As String, Optional ByVal createIfMissing = False)
    With Config.DropDowns(dropDownBoxName)
        If .List(.ListIndex) = "(Automatic)" Then
            Dim colName As String
            colName = Mid(dropDownBoxName, 3) 'i.e. If list box is called ddLattitude, name is Lattitude
            'Verify that the column doesn't exist, if it does, use it.
            getColumnFor = getColumnNumber(colName)
            If getColumnFor = 0 Then 'The column doesn't exist
                If createIfMissing Then
                        wkstLocations.Cells(1, wkstLocations.UsedRange.Columns.Count + 1).value = colName
                        DataRangesRoutines.UpdateRanges
                Else 'We aren't creating the new column, one doesn't exist, so return 0
                    getColumnFor = 0
                    Exit Function
                End If
            Else 'The column exists, use it!
                .ListIndex = getColumnFor
            End If
        End If
        getColumnFor = .ListIndex
    End With
End Function
