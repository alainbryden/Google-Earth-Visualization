VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wkstLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public somethingChanged As Boolean
Private sheetActive As Boolean

'Runs whenever the worksheet is activated
Private Sub Worksheet_Activate()
    If ActiveSheet.Name <> Me.Name Then Exit Sub
    Dim rememberactive As Range
    If colUniqueID() = 0 Then Exit Sub
    Set rememberactive = Application.ActiveCell
    Application.ScreenUpdating = False
    'This section refreshes conditional formatting that highlights rows with no lat/long info on the sheet
    With wkstLocations
        .Cells(1, 1).Activate
        .Cells.FormatConditions.Delete
        If .UsedRange.Rows.Count > 3 And colLattitude() <> 0 And colLongitude() <> 0 Then
            With .Range(.Cells(1, colUniqueID()), .Cells(.UsedRange.Rows.Count, colUniqueID()))
                .FormatConditions.Add Type:=xlExpression, _
                    Formula1:="=OR(ISBLANK(" & wkstLocations.Cells(1, colLattitude()).Address(False, True) & ")," _
                                & "ISBLANK(" & wkstLocations.Cells(1, colLongitude()).Address(False, True) & "))"
                .FormatConditions(1).Interior.color = RGB(220, 150, 150)
            End With
        End If
    End With
    rememberactive.Activate 'Put the user's cursor back to the last place it was
    sheetActive = True
    somethingChanged = False
    Application.ScreenUpdating = True
End Sub

'Detects when something in the exposure sheet is changed, so that the changes can be processed
Private Sub Worksheet_Change(ByVal Target As Range)
    If ActiveSheet.Name <> Me.Name Then Exit Sub
    If sheetActive Then somethingChanged = True
End Sub

'If something was changed in the exposures, reprocesses the data upon leaving the sheet
Private Sub Worksheet_Deactivate()
    If somethingChanged Then
        Automation True
        Application.StatusBar = "Analysing Data"
        DataRangesRoutines.UpdateRanges
        ConfigControlsRoutines.UpdateControls
        ToolsControlsRoutines.UpdateControls
        Application.StatusBar = False
        Automation False
    End If
    sheetActive = False
End Sub
