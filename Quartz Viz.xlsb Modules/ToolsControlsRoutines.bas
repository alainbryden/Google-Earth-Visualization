Attribute VB_Name = "ToolsControlsRoutines"
Option Explicit

'Names of named ranges in the Tools worksheet
Public Const rngGeocoderIDrow = "GeocoderSelectedID"
Public Const rngLocatorIDrow = "LocatorSelectedID"
Public Const rngPlotSelectedCategs = "PlotSelectedCategs"
Public Const rngColourScheme = "ColourSchemeTable"
'The range where various useful strings are contained
Public Const rngGeocoderAddress = "GeocoderAddress"
Public Const rngGeocoderDescription = "GeocoderDescription"
Public Const rngLocatorDescription = "LocatorDescription"
'Names of controls in the Tools worksheet
Public Const ctrl_ddGeocoderID = "ddGeocoderID"
Public Const ctrl_ddLocatorID = "ddLocatorID"
Public Const ctrl_ddLocatorCateg = "ddLocatorCateg"
Public Const ctrl_ddLocatorSelectedCateg = "ddLocatorSelectedCateg"

Public Const ctrl_PlotGroupBox = "PlotGroupBox"
Public Const ctrl_UnitSlider = "UnitSlider"
Public Const ctrl_optRecycleData = "optRecycleData"
Public Const ctrl_ddPlotValue = "ddPlotValue"
Public Const ctrl_ddPlotCateg = "ddPlotCateg"
Public Const ctrl_optBarGraph = "optBarGraph"
Public Const ctrl_optSum = "optSum"
Public Const ctrl_optAverage = "optAverage"
Public Const ctrl_optThematicMap = "optThematicMap"
Public Const ctrl_opt3DBars = "opt3DBars"

Const PlotGroupDefaultHeight = 185

Private prevddPlotCategVal As Variant
Private prevPlotValueVal As Variant
Private prevUnitSliderVal As Variant
Private prevVisualizationTypeVal As Variant
Private prevAggTypeVal As Variant

'Remembers the state of the controls so that a change can be detected
Public Sub RememberControlSettings()
    prevUnitSliderVal = UnitSlider().value
    prevddPlotCategVal = ddPlotCateg().value
    prevPlotValueVal = ddPlotValue().value
    prevVisualizationTypeVal = optThematicMap().value
    prevAggTypeVal = optSum().value
End Sub

Public Sub UpdateControls()
    Automation True
    'Unprotect the sheet so that elements can be edited
    Dim prevProtection: prevProtection = Tools.ProtectionMode: Tools.Unprotect
    'Get the unique identifier captions
    Dim uniqueIDtext As String
    uniqueIDtext = wkstLocations.Cells(1, colUniqueID).Text
    With lblGeocoderID()
        .Caption = IIf(Len(uniqueIDtext) = 0, "Unique ID", uniqueIDtext) & ":"
        While .Width > 75
            .Caption = Mid(.Caption, 1, Len(.Caption) - 5) & "...:"
        Wend
        .Left = 80 - .Width
    End With
    With lblLocatorID()
        .Caption = lblGeocoderID().Caption
        .Left = lblGeocoderID().Left
    End With
    'Refresh the contents of the unique identifier drop down boxes
    RefreshUniqueIdentifierList ddLocatorID()
    RefreshUniqueIdentifierList ddGeocoderID()
    'Refresh the lists
    ddPlotCateg().ListFillRange = rngAvailableCategories
    ddLocatorCateg().ListFillRange = rngAvailableCategories
    'Reset the values of controls where the value is invalid
    If ddPlotValue().ListIndex = 0 Then ddPlotValue().ListIndex = 1
    If ddPlotCateg().ListIndex = 0 Then ddPlotCateg().ListIndex = 1
    If ddLocatorCateg().ListIndex = 0 Then ddLocatorCateg().ListIndex = 1
    If ddLocatorSelectedCateg().ListIndex = 0 Then ddLocatorSelectedCateg().ListIndex = 1
    If Not dataRecalled Then
        optRecycleData() = xlOff
        OptRecycleDataEnabled False
    End If
    updateMaxBarHeight
    If prevProtection = True Then Tools.Protect

    Automation False
End Sub

'Refresh the contents of the unique identifier drop down boxes
Private Function RefreshUniqueIdentifierList(ByRef control As DropDown)
    Dim prevIndex As Integer
    With control
        prevIndex = .ListIndex
        .ListFillRange = ""
        If (wkstLocations.UsedRange.Rows.Count < 2) Then
            .RemoveAllItems
            .AddItem ("No Data Present")
            .ListIndex = 1
        Else
            .ListFillRange = "'" & wkstLocations.Name & "'!" & toEndOfColumn(wkstLocations.Cells(2, colUniqueID())).Address
            If prevIndex <= .listCount Then .ListIndex = prevIndex
        End If
    End With
End Function

Public Function validateControls() As Boolean
    'Check that required values are specified in the exposures.
    validateControls = True
    If colLattitude() = 0 Or colLongitude() = 0 Then
        MsgBox "Cannot display locations because Lattitude and Longitude are not specified anywhere in the worksheet." _
              & " To configure the location, go to the Advanced Configuration sheet (hidden by default).", vbError
        validateControls = False
    End If
    If getSelectedPlotValueCol() = 0 Then
        MsgBox "Cannot proceed because you have not picked a value to plot.", vbError
        validateControls = False
    End If
    Dim cel As Range
End Function

'Functions for returning the control objects in the spreadsheet.
Public Function lblGeocoderID() As Object
    Set lblGeocoderID = Tools.GeocoderUniqueIDLabel
End Function
Public Function lblLocatorID() As Object
    Set lblLocatorID = Tools.LocatorUniqueIDLabel
End Function
Public Function lblGeoedLat() As Object
    Set lblGeoedLat = Tools.GeoedLat
End Function
Public Function lblGeoedLong() As Object
    Set lblGeoedLong = Tools.GeoedLong
End Function
Public Function ddGeocoderID() As DropDown
    Set ddGeocoderID = Tools.DropDowns(ctrl_ddGeocoderID)
End Function
Public Function ddLocatorID() As DropDown
    Set ddLocatorID = Tools.DropDowns(ctrl_ddLocatorID)
End Function
Public Function ddLocatorCateg() As DropDown
    Set ddLocatorCateg = Tools.DropDowns(ctrl_ddLocatorCateg)
End Function
Public Function ddLocatorSelectedCateg() As DropDown
    Set ddLocatorSelectedCateg = Tools.DropDowns(ctrl_ddLocatorSelectedCateg)
End Function
Public Function PlotGroupBox() As GroupBox
    Set PlotGroupBox = Tools.GroupBoxes(ctrl_PlotGroupBox)
End Function
Public Function UnitSlider() As ScrollBar
    Set UnitSlider = Tools.ScrollBars(ctrl_UnitSlider)
End Function
Public Function optRecycleData() As CheckBox
    Set optRecycleData = Tools.CheckBoxes(ctrl_optRecycleData)
End Function
Public Function optSum() As OptionButton
    Set optSum = Tools.OptionButtons(ctrl_optSum)
End Function
Public Function optAverage() As OptionButton
    Set optAverage = Tools.OptionButtons(ctrl_optAverage)
End Function
Public Function ddPlotValue() As DropDown
    Set ddPlotValue = Tools.DropDowns(ctrl_ddPlotValue)
End Function
Public Function ddPlotCateg() As DropDown
    Set ddPlotCateg = Tools.DropDowns(ctrl_ddPlotCateg)
End Function
Public Function optBarGraph() As CheckBox
    Set optBarGraph = Tools.CheckBoxes(ctrl_optBarGraph)
End Function
Public Function optThematicMap() As CheckBox
    Set optThematicMap = Tools.CheckBoxes(ctrl_optThematicMap)
End Function
Public Function opt3DBars() As CheckBox
    Set opt3DBars = Tools.CheckBoxes(ctrl_opt3DBars)
End Function

'Get the column number containing the value selected to be plotted
Public Function getSelectedPlotValueCol() As Integer
    With ddPlotValue()
        getSelectedPlotValueCol = getColumnNumber(.List(.ListIndex))
    End With
End Function

'Get the column number containing the value selected to be plotted
Public Function getSelectedPlotCategoryCol() As Integer
    With ddPlotCateg()
        getSelectedPlotCategoryCol = getColumnNumber(.List(.ListIndex))
    End With
End Function

'Updates the Tools.Range in the tools spreadsheet (under this control) containing which row is selected
Private Sub ddGeocoderID_Change()
    Const FormulaAddr = "=INDIRECT(ADDRESS(GeocoderSelectedID,rngColLookupAddress,,,""Exposure Information""))"
    Const emptyString = "No look-up address string specified. See ""Advanced Configuration"" sheet (hidden by default)."
    Static prevVal As Variant 'used to prevent this event from running unecessarily
    Dim AddressString As String
    If ddGeocoderID().value = prevVal Then Exit Sub
    prevVal = ddGeocoderID().value
    Dim prevProtection: prevProtection = Tools.ProtectionMode: Tools.Unprotect
        Tools.Range(rngGeocoderIDrow).value = ddGeocoderID().ListIndex + 1
        With Tools.Range(rngGeocoderAddress)
            If ConfigControlsRoutines.colLookupAddress() = 0 Then
                .value = emptyString
            Else
                .Formula = FormulaAddr
                .Calculate
                .Formula = .value
            End If
        End With
        lblGeoedLat().Caption = ""
        lblGeoedLong().Caption = ""
    If prevProtection = True Then Tools.Protect
End Sub

'Updates the Tools.Range in the tools spreadsheet (under this control) containing which row is selected
Private Sub ddLocatorID_Change()
    Const emptyString = "No description string specified. See ""Advanced Configuration"" sheet (hidden by default)."
    Static prevVal As Variant 'used to prevent this event from running unecessarily
    Dim DescriptionString As String
    If ddLocatorID().value = prevVal Then Exit Sub
    prevVal = ddLocatorID().value
    Dim prevProtection: prevProtection = Tools.ProtectionMode: Tools.Unprotect
    Tools.Range(rngLocatorIDrow).value = ddLocatorID().ListIndex + 1
    If prevProtection = True Then Tools.Protect
End Sub

'Control allowing the user to view locations matching a certain category in a categorization
Private Sub ddLocatorCateg_Change()
    Static prevVal As Variant 'used to prevent this event from running unecessarily
    If ddLocatorCateg().value = prevVal Then Exit Sub
    prevVal = ddLocatorCateg().value
    Application.Calculation = xlCalculationManual
    DataRangesRoutines.RefreshViewLocationsCategories ddLocatorCateg().List(ddLocatorCateg().ListIndex)
    With ddLocatorSelectedCateg()
        .ListFillRange = rngLocatorCategories
        .enabled = (ddLocatorCateg().ListIndex <> 1)
        .ListIndex = 1
    End With
    Application.Calculation = xlCalculationAutomatic
End Sub

'Add a row to the bar graph goupings colour schemes
Public Function btnAddGrouping_Click()
    Dim categoryCount As Integer
    Dim rowStart As Integer, colStart As Integer, rowEnd As Integer, colEnd As Integer
    With Tools.Range(rngColourScheme)
        rowStart = .row: colStart = .column
        categoryCount = .Rows.Count
        rowEnd = .row + .Rows.Count - 1
        colEnd = .column + .Columns.Count - 1
    End With
    With Tools
        PlotGroupBox().Height = Application.WorksheetFunction.Max(PlotGroupDefaultHeight, _
                                (Tools.Range(rngPlotSelectedCategs)(1).RowHeight) * (categoryCount + 5))
        .Range(.Cells(rowEnd, colStart), .Cells(rowEnd, colEnd)).Copy .Cells(rowEnd + 1, colStart)
        Application.CutCopyMode = False
        Names(rngColourScheme).RefersTo = .Range(.Cells(rowStart, colStart), .Cells(rowEnd + 1, colEnd))
        Names(rngPlotSelectedCategs).RefersTo = .Range(.Cells(rowStart, colStart), .Cells(rowEnd + 1, colStart))
    End With
End Function

'Add a row to the bar graph goupings colour schemes
Public Function btnRemoveGrouping_Click()
    Dim categoryCount As Integer
    Dim rowStart As Integer, colStart As Integer, rowEnd As Integer, colEnd As Integer
    With Tools.Range(rngColourScheme)
        rowStart = .row: colStart = .column
        categoryCount = .Rows.Count
        rowEnd = .row + .Rows.Count - 1
        colEnd = .column + .Columns.Count - 1
    End With
    If categoryCount = 1 Then
        MsgBox "If you wish to not break locations into categories, change the last category to (All Remaining Locations) " _
               & "or change the visualization type to 'Thematic Map'", vbOKOnly, "Cannot delete last category"
        Exit Function
    End If
    With Tools
        PlotGroupBox().Height = Application.WorksheetFunction.Max(PlotGroupDefaultHeight, _
                                (Tools.Range(rngPlotSelectedCategs)(1).RowHeight) * (categoryCount + 3))
        .Range(.Cells(rowEnd, colStart), .Cells(rowEnd, colEnd)).Delete xlUp
        Names(rngColourScheme).RefersTo = .Range(.Cells(rowStart, colStart), .Cells(rowEnd - 1, colEnd))
        Names(rngPlotSelectedCategs).RefersTo = .Range(.Cells(rowStart, colStart), .Cells(rowEnd - 1, colStart))
    End With
End Function

Public Sub UnitSlider_Change()
    If optRecycleData().enabled Then If Not (UnitSlider().value = prevUnitSliderVal) Then optRecycleData() = xlOff
    If Tools.lblUnitSliderCurValue.Visible = False Then Tools.lblUnitSliderCurValue.Visible = True
    Tools.lblUnitSliderCurValue.Caption = sigFigs(getUnitValue * 100, 4)
    updateMaxBarHeight
End Sub

'Get the grouping size by interpreting the slider value on a logarithmic scale
Public Function getUnitValue()
    getUnitValue = SlderToGroupSize(UnitSlider().value)
End Function

Private Function SlderToGroupSize(ByVal sliderValue As Double) As Double
    Const SliderMax = 30000, SliderMin = 1, AreaMax = 10, AreaMin = 0.001
    SlderToGroupSize = (AreaMax / AreaMin) ^ ((sliderValue - SliderMin) / (SliderMax - SliderMin)) * AreaMin
End Function

Private Function GroupSizeToSlider(ByVal groupSize As Double) As Double
    Const SliderMax = 30000, SliderMin = 1, AreaMax = 10, AreaMin = 0.001
    GroupSizeToSlider = (Log(groupSize / AreaMin) / Log(AreaMax / AreaMin)) * (SliderMax - SliderMin) + SliderMin
End Function


'Lets the user plot data differently using previously generated data (faster)
Public Sub optRecycleData_Click()
    If Not dataRecalled Then
        optRecycleData() = xlOff
        OptRecycleDataEnabled False
    End If
    If optRecycleData() = xlOn Then
        UnitSlider() = prevUnitSliderVal
        ddPlotCateg() = prevddPlotCategVal
        ddPlotValue() = prevPlotValueVal
        optSum() = prevAggTypeVal
        optAverage() = IIf(prevAggTypeVal = xlOn, xlOff, xlOn)
        setOptThematicMapChecked IIf(prevVisualizationTypeVal = xlOn, True, False)
    End If
End Sub

Public Sub optSum_Change()
    If optRecycleData().enabled Then If Not (optSum() = prevAggTypeVal) Then optRecycleData().value = False
End Sub
Public Sub optAverage_Click()
    optSum_Change
End Sub

Public Sub ddPlotValue_Change()
    Static prevVal As Variant 'used to prevent this event from running unecessarily
    If ddPlotValue().value = prevVal Then Exit Sub
    prevVal = ddPlotValue().value
    If optRecycleData().enabled Then If Not (ddPlotValue().value = prevPlotValueVal) Then optRecycleData().value = False
End Sub

Public Sub ddPlotCateg_Change()
    Static prevVal As Variant 'used to prevent this event from running unecessarily
    If ddPlotCateg().value = prevVal Then Exit Sub
    prevVal = ddPlotCateg().value
    If optRecycleData().enabled Then If Not (ddPlotCateg().value = prevddPlotCategVal) Then optRecycleData().value = False
    Application.Calculation = xlCalculationManual
    DataRangesRoutines.RefreshPlotterCategories ddPlotCateg().List(ddPlotCateg.ListIndex)
    setOptThematicMapChecked (ddPlotCateg().ListIndex = 1)
    Application.Calculation = xlCalculationAutomatic
End Sub

'Lets the user select the bar graph visualization
Public Sub optBarGraph_Click()
    setOptThematicMapChecked optBarGraph().value = xlOff
End Sub

Public Sub setOptThematicMapChecked(ByVal checked As Boolean)
    optThematicMap().value = IIf(checked, xlOn, xlOff)
    optThematicMap_Click 'unfortunately, must be run manually when changed programatically
End Sub

'Lets the user select whether they wish to create a thematic map
'This routine is responsible for ensuring that the appropriate relationship is maintained between
'the controls in the graph section. Some controls are enabled when others are checked etc.
'It used to be that this ran automatically whenever optthematicmap value changed, but as a form control,
'this doesn't happen, so everytime optThematicMap's value is changed programatically, this method MUST be
'called afterwards.
Public Sub optThematicMap_Click()
    Dim optChecked As Boolean
    optChecked = IIf(optThematicMap().value = xlOn, True, False)
    With optThematicMap()
        optBarGraph().value = IIf(optChecked, xlOff, xlOn)
        Opt3DBarsEnabled optChecked
        If (optThematicMap().value <> prevVisualizationTypeVal) Then optRecycleData() = xlOff
        Tools.Range(rngColourScheme).Locked = optChecked 'Lock and grey out colour scheme if using thematic map
        Tools.Range(rngColourScheme).Font.color = IIf(optChecked, RGB(200, 200, 200), RGB(0, 0, 0))
        If optChecked Then
            ddPlotCateg().ListIndex = 1
        Else
            setOpt3DBarsChecked True
            If ddPlotCateg().ListIndex = 1 Then If ddPlotCateg().listCount > 1 Then ddPlotCateg().ListIndex() = 2
        End If
    End With
End Sub

'Lets the user choose to use 3D bars or not
Public Sub setOpt3DBarsChecked(ByVal checked As Boolean)
    opt3DBars().value = IIf(checked, xlOn, xlOff)
    opt3DBars_Click
End Sub

Public Sub opt3DBars_Click()
    updateMaxBarHeight
End Sub

Public Sub updateMaxBarHeight()
    Dim optChecked As Boolean
    optChecked = IIf(opt3DBars().value = xlOn, True, False)
    Tools.lblMaxBarHeight.Visible = optChecked
    With Tools.MaxBarHeight
        .Visible = optChecked
        If (Not IsNumeric(.Text)) Then .Text = "Auto (" & sigFigs(autoMaxBarHeight(), 3) & ")"
    End With
End Sub

'Determine an auto max height based on the location grouping size
Public Function autoMaxBarHeight(Optional ByVal unit As Double)
    If unit = 0 Then unit = getUnitValue()
    autoMaxBarHeight = 10 ^ 2 * unit
End Function


'Routine that works with z order and hidden controls to give checkboxes the appearance of being disabled
Public Sub OptRecycleDataEnabled(ByVal enabled As Boolean)
    optRecycleData().enabled = enabled
    Tools.Shapes(ctrl_optRecycleData).ZOrder IIf(enabled, msoBringToFront, msoSendToBack)
End Sub

'Routine that works with z order and hidden controls to give checkboxes the appearance of being disabled
Public Sub Opt3DBarsEnabled(ByVal enabled As Boolean)
    opt3DBars().enabled = enabled
    setOpt3DBarsChecked enabled
    Tools.Shapes(ctrl_opt3DBars).ZOrder IIf(enabled, msoBringToFront, msoSendToBack)
End Sub

