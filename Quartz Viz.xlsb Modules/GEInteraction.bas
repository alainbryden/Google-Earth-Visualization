Attribute VB_Name = "GEInteraction"
Option Explicit

#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'Data that the user can store to quickly generate different plot styles
Global dataRecalled As Boolean
Dim Buckets As Map3D
'Dim gridSum() As Double, bucketSize() As Long 'The value bucket and its item count
Dim maxValue As Double 'The maximum value found when using the thematic map

Private Function GEInit() As ApplicationGE
    'Start google earth and wait for the application to initialize
    Application.StatusBar = "Initializing Google Earth..."
    On Error GoTo noGE
    Set GEInit = CreateObject("GoogleEarth.ApplicationGE")
    While (GEInit.IsInitialized = 0): DoEvents: Wend
    GoTo finally
noGE:
    MsgBox "Google earth is either not properly installed on this machine or isn't responding.", vbCritical
    Set GEInit = Nothing
finally:
    Application.StatusBar = False
End Function

Public Sub LookUpAddress_Click()
    Dim GEI As ApplicationGE
    Dim PointOnTerrain As PointOnTerrainGE
    Dim Search As SearchControllerGE
    Dim KMLData As String
    Dim row As Long
    
    'Get the row of the location selected to look up
    With ddGeocoderID()
        If .listCount <= 1 Then Exit Sub
        row = .ListIndex + 1
    End With
    
    Set GEI = GEInit()
    If GEI Is Nothing Then Exit Sub
    
    KMLData = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
            "<kml xmlns=""http://www.opengis.net/kml/2.2"">" & _
            "<Placemark>" & _
                "<name>" & ddGeocoderID().List(ddGeocoderID().ListIndex) & "</name>" & _
                "<visibility>1</visibility>" & _
                "<open>1</open>" & _
                "<description>" & "<![CDATA[" & Tools.Range(rngGeocoderDescription) & "]]></description>" & _
                "<address>" & Tools.Range(rngGeocoderAddress) & "</address>" & _
            "</Placemark>" & _
            "</kml>"
    GEI.LoadKmlData KMLData
    
    Set Search = GEI.SearchController()
    Call Search.Search(Tools.Range(rngGeocoderAddress))
    Dim resul As Variant
    Set resul = Search.GetResults()
    Dim lat As Double, lon As Double, prevlat As Double, prevlon As Double, checkChange As Double
    Dim steady As Boolean
    steady = False: checkChange = -1
    lat = 0: lon = 0: prevlat = -1: prevlon = -1
    While Not steady
        Set PointOnTerrain = GEI.GetPointOnTerrainFromScreenCoords(0, 0)
        lat = PointOnTerrain.Latitude
        lon = PointOnTerrain.Longitude
        lblGeoedLat().Caption = sigFigs(lat, 8)
        lblGeoedLong().Caption = sigFigs(lon, 8)
        DoEvents
        If (checkChange = 100) Then
            If (lat = prevlat And lon = prevlon) Then steady = True
            prevlat = lat: prevlon = lon
            checkChange = -1
        End If
        checkChange = checkChange + 1
        Sleep 10
    Wend
End Sub

'Takes the Lat/Long read from google earth and write it to the spreadsheet
Public Sub RecordNewLatLong_Click()
    Dim row As Long
    row = ddGeocoderID().ListIndex + 1
    'Write the values. Create a lat and long column if they were configured for "(Automatic)"
    wkstLocations.Cells(row, colLattitude(True)).value = lblGeoedLat().Caption
    wkstLocations.Cells(row, colLongitude(True)).value = lblGeoedLong().Caption
End Sub

Public Sub GetCurrentLocation_Click()
    Dim GEI As ApplicationGE
    Dim PointOnTerrain As PointOnTerrainGE
    Set GEI = GEInit()
    If GEI Is Nothing Then Exit Sub
    Set PointOnTerrain = GEI.GetPointOnTerrainFromScreenCoords(0, 0)
    lblGeoedLat().Caption = sigFigs(PointOnTerrain.Latitude, 8)
    lblGeoedLong().Caption = sigFigs(PointOnTerrain.Longitude, 8)
End Sub

Public Sub ViewAllLocations_Click()
    Dim GEI As ApplicationGE
    Dim lon As Double, lat As Double
    Dim colLat As Integer, colLon As Integer, colID As Integer, colCategory As Integer, colDescription As Integer
    Dim i As Long, lastRow As Long, noData As Long, noMatch As Long
    Dim KMLData As String
    Dim uniqueIdentifier As String, categoryToMatch As String
    
    'Get some numbers used a lot during the loop into a variable
    colLat = ConfigControlsRoutines.colLattitude()
    colLon = ConfigControlsRoutines.colLongitude()
    colID = ConfigControlsRoutines.colUniqueID()
    colDescription = ConfigControlsRoutines.colDescription()
    If ddLocatorSelectedCateg().ListIndex = 1 Then
        colCategory = 0
    Else
        colCategory = getColumnNumber(ddLocatorCateg().List(ddLocatorCateg().ListIndex))
    End If
    categoryToMatch = ddLocatorSelectedCateg().List(ddLocatorSelectedCateg().ListIndex)
    If categoryToMatch = "(Blank)" Then categoryToMatch = ""
    
    'Check that lattitude and longitude are specified in the exposures.
    If colLat = 0 Or colLon = 0 Then
        MsgBox "Cannot display locations because Lattitude and Longitude are not specified anywhere in the worksheet." _
              & " To configure the location, go to the Advanced Configuration sheet (hidden by default).", vbError
        Exit Sub
    End If
    
    Set GEI = GEInit
    If GEI Is Nothing Then Exit Sub
    
    'Will use a special concatenation routine for building large strings
    Dim catOffset As Long: catOffset = 0
    'Assign a data block to hold the kml data
    KMLData = String(300, " ")
    
    fastConcat KMLData, "<?xml version=""1.0"" encoding=""UTF-8""?><kml xmlns=""http://www.opengis.net/kml/2.2"">" & _
                        "<Document id=""AllLocations""><name>" & xmlFormatString(ddLocatorSelectedCateg().List(ddLocatorSelectedCateg().ListIndex)) & "</name>", catOffset
    'Write each row of data with a lat/long into the kml data to send to Google Earth
    With wkstLocations
        noData = 0: noMatch = 0
        lastRow = endOfColumn(wkstLocations, colID).row
        For i = 2 To lastRow
            If lastRow < 1000 Or i Mod 100 = 1 Then
                Application.StatusBar = "Compiling Locations: " & (i - 1) & "/" & (lastRow - 1) & _
                                "(" & Int(10000 * ((i - 1) / (lastRow - 1))) / 100 & "%) complete..."
                DoEvents
            End If
            Dim validLocation As Boolean
            validLocation = (colCategory = 0)
            If Not validLocation Then validLocation = (.Cells(i, colCategory).value = categoryToMatch)
            If validLocation Then
                If .Cells(i, colLat).value <> Empty And .Cells(i, colLon).value <> Empty Then
                    uniqueIdentifier = xmlFormatString(.Cells(i, colID))
                    lat = CDbl(.Cells(i, colLat))
                    lon = CDbl(.Cells(i, colLon))
                    fastConcat KMLData, _
                        "<Placemark>" & _
                            "<name>" & uniqueIdentifier & "</name>" & _
                            "<visibility>1</visibility>" & _
                            "<description>" & "<![CDATA[" & .Cells(i, colDescription) & "]]></description>" & _
                            "<Point>" & _
                                "<coordinates>" & lon & "," & lat & ",0" & "</coordinates>" & _
                            "</Point>" & _
                        "</Placemark>", catOffset
                Else
                    noData = noData + 1
                End If
            Else
                noMatch = noMatch + 1
            End If
        Next i
    End With
    wkstLocations.AutoFilterMode = False
    Application.StatusBar = "Pushing " & (lastRow - noData - noMatch - 1) & " of " & (lastRow - 1) & _
                            " Locations to Google Earth (" & _
                            noMatch & " locations didn't match, " & noData & " locations had no lat/long)"
    fastConcat KMLData, "</Document></kml>", catOffset
    On Error GoTo tryfile
    GEI.LoadKmlData Trim(KMLData) 'Send the kml data string to google earth
    Application.StatusBar = False
    Exit Sub
tryfile: 'Try writing data to a temporary file if the other method failed.
    If MsgBox("There is too much data, and Google Earth won't load it directly. Do you want to try to " _
            & "write data to a temporary file and load it?", vbYesNo, "Too much data") = vbYes Then
        Dim Fnum As Integer
        Dim Fname As String
        Fnum = FreeFile()
        Fname = "C:\Temp\AllLocations.kml"
        Open Fname For Output As #Fnum
        Print #Fnum, KMLData
        Close #Fnum
        Shell (CStr("C:\Program Files\Google\Google Earth\googleearth.exe") & " " & Fname)
        MsgBox "Hopefully that worked. The file is at " & Fname & ", delete it when finished", , "Gave it my best shot"
    End If
    Application.StatusBar = False
End Sub

Public Sub ViewLocation_Click()
    Dim lon As Double, lat As Double
    Dim row As Integer
    Dim uniqueIdentifier As String
    Dim KMLData As String
    Dim GEI As ApplicationGE
    Dim Feature As FeatureGE
    
    'Get the row that we are about to show in Google Earth
    If Tools.Range(rngLocatorIDrow) <= 1 Then Exit Sub
    row = Tools.Range(rngLocatorIDrow)
    uniqueIdentifier = wkstLocations.Cells(row, colUniqueID())
    
    'Check that lattitude and longitude are specified in the exposures.
    If colLattitude() = 0 Or colLongitude() = 0 Then
        MsgBox "Cannot display locations because Lattitude and Longitude are not specified anywhere in the worksheet." _
              & " To configure the location, go to the Advanced Configuration sheet (hidden by default).", vbError
        Exit Sub
    End If
    
    'Check that a lat and long are specified for this row
    If IsEmpty(wkstLocations.Cells(row, colLattitude()).value) Or IsEmpty(wkstLocations.Cells(row, colLongitude()).value) Then
        MsgBox ("Lattitude and/or Longitude is not specified for selected Row")
        Exit Sub
    End If
    
    'Get the lattitude and longitude of the current location
    lat = wkstLocations.Cells(row, colLattitude())
    lon = wkstLocations.Cells(row, colLongitude())
    
    Set GEI = GEInit
    If GEI Is Nothing Then Exit Sub
    
    'Create the KML string to send directly to the Google Earth application
    KMLData = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
            "<kml xmlns=""http://www.opengis.net/kml/2.2"">" & _
            "<Placemark>" & _
                "<name>" & xmlFormatString(uniqueIdentifier) & "</name>" & _
                "<visibility>1</visibility>" & _
                "<open>1</open>" & _
                "<description>" & "<![CDATA[" & Tools.Range(rngLocatorDescription) & "]]></description>" & _
                "<Point>" & _
                    "<coordinates>" & lon & "," & lat & ",0" & "</coordinates>" & _
                "</Point>" & _
            "</Placemark>" & _
            "</kml>"
    GEI.LoadKmlData KMLData
    Set Feature = GEI.GetFeatureByName(uniqueIdentifier)
    On Error Resume Next 'Might not succeed if multiple features have the same name
    Call Feature.Highlight
    Call GEI.SetFeatureView(Feature, 0.5)
End Sub

Public Sub GenerateVisualization_Click()
    Dim GEI As ApplicationGE
    Dim col As Long, row As Long, bar As Integer 'Loop counters for region and category groupings
    Dim i As Long, totalRows As Long 'Temp variable for kml row number, and total rows
    Dim minLat As Double, minLong As Double 'Stores the min lat/long from the locations sheet
    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double 'used to store temporary co-ordinates
    Dim unit As Double, Height As Double, MaxBarHeight As Double
    Dim numbars As Integer 'number of categorizations
    Dim KMLData As String
    Dim categorizations As Boolean: categorizations = (optThematicMap() = xlOff)
    Dim vizName As String
    Dim aggType As Integer
    If optSum() = xlOn Then aggType = 1 Else If optAverage() = xlOn Then aggType = 2
    
    'Additional variables for bar graph only
    Dim color() As String, barType() As String 'The categories and colours to use in bar graph
    Dim tempDescStr As String
    Dim putGrid As Boolean
    Dim currBar As Integer, barsHere As Integer
    
    If Not validateControls() Then Exit Sub 'Check that selected controls are valid
    Set GEI = GEInit
    If GEI Is Nothing Then Exit Sub
    Automation True
    
    'Strings for the kml file
    Dim formatStr As String, valueName As String, altitudeMode As String
    
    'Determine the the proper format string
    With endOfColumn(wkstLocations, getSelectedPlotValueCol())
        formatStr = .NumberFormat 'Get the number format from the column containing the values
        If formatStr = "General" Then formatStr = "General Number" 'Or an overflow error will occur
        'Test if the format is compatible with the format function. If not, compile a default format.
        If Not (.Text = Format(.value, formatStr)) Then
            formatStr = "#,###"
            If InStr(1, .Text, ".") <> 0 Then formatStr = formatStr & ".0######"
            If InStr(1, .Text, "$") <> 0 Then formatStr = "$" & formatStr
            If InStr(1, .Text, "%") <> 0 Then formatStr = formatStr & "%"
        End If
    End With
    
    valueName = xmlFormatString(wkstLocations.Cells(1, getSelectedPlotValueCol()).Text) 'Get the title of the values from the column name
    altitudeMode = IIf(opt3DBars().value = xlOn, "relativeToGround", "snapToGround")
    
    Application.StatusBar = "Generating KML Data..."
    Dim catOffset As Long: catOffset = 0 'Will use a special concatenation routine for building large strings
    KMLData = String(1000, " ") 'Assign a large data block to hold the kml data
    vizName = IIf(categorizations, "Bar Graph", "Thematic Map") & " of " _
              & IIf(aggType = 1, "Total ", "Average ") & ddPlotValue.List(ddPlotValue.ListIndex)
    'Begin the KML document for the visualization
    fastConcat KMLData, "<?xml version=""1.0"" encoding=""UTF-8""?><kml xmlns=""http://www.opengis.net/kml/2.2"">" _
                      & "<Document>" & "<name>" & vizName & "</name>", catOffset
                   
    Application.StatusBar = "Gathering Required Data..."
    If categorizations Then
        'Get the colour scheme defined
        Dim cel As Range, hexColor As String
        ReDim barType(0 To Tools.Range(rngPlotSelectedCategs).Count - 1)
        ReDim color(LBound(barType) To UBound(barType))
        i = LBound(barType)
        For Each cel In Tools.Range(rngPlotSelectedCategs)
            barType(i) = cel.Text
            hexColor = Hex(ConditionalColor(cel.Offset(0, 2)))
            If Len(hexColor) < 6 Then hexColor = String(6 - Len(hexColor), "0") & hexColor
            color(i) = "#" & Hex(cel.Offset(0, 3).value * 255) & hexColor
            i = i + 1
        Next cel
        putGrid = False
        'Define the colours that might be used in the bar graphs
        fastConcat KMLData, strColorStyles(color), catOffset
    End If
    
    'Get other values needed for plotting
    minLat = Application.WorksheetFunction.Min(toEndOfColumn(wkstLocations.Cells(2, colLattitude())))
    minLong = Application.WorksheetFunction.Min(toEndOfColumn(wkstLocations.Cells(2, colLongitude())))
    
    'Compute the maximum bar height for 3D plots
    unit = ToolsControlsRoutines.getUnitValue()
    If IsNumeric(Tools.MaxBarHeight.Text) Then
        MaxBarHeight = Val(Tools.MaxBarHeight.Text) * 10000
    Else
        MaxBarHeight = autoMaxBarHeight(unit) * 10000
    End If
    
    'Data for plotting
    ToolsControlsRoutines.optRecycleData_Click 'Make sure that if 'use previously gathered data' is selected, the data is in memory.
    If categorizations Then 'Get Bar Graph
        numbars = UBound(barType) - LBound(barType) + 1
        If optRecycleData().value = xlOff Then retrieveLocationsData True, barType
    Else 'Get Thematic Map data
        numbars = 1
        If optRecycleData().value = xlOff Then retrieveLocationsData False
    End If
    If maxValue = 0 Then maxValue = 1
        
    totalRows = Buckets.Count
    Dim arrayData(), bucket As DataBucket, bucketValue As Double
    Application.StatusBar = "Sorting locations for optimal rendering."
    arrayData = Buckets.toArray()
        
    'Begin the main loop for generating kml data for each region subdivision
    'Perform in reverse order so that when looking towards the north and east, elements ahead were
    'rendered first and so can be seen through transparent surfaces
    For i = UBound(arrayData) To LBound(arrayData) Step -1
        Set bucket = arrayData(i)
        
        'Initialize the data for a cell (x,y) of bars (z) if we are in a new cell
        If categorizations And (col <> bucket.j Or row <> bucket.i) Then
            tempDescStr = IIf(aggType = 1, "Total ", "Average ") & valueName & "<br/>"
            currBar = 1
        End If
        
        row = bucket.i
        col = bucket.j
        bar = bucket.k
        If categorizations Then barsHere = Buckets.Items(row, col).Count
        If aggType = 1 Then bucketValue = bucket.Sum Else bucketValue = bucket.Average
        

        If totalRows < 1000 Or i Mod 100 = 1 Then 'Update the status bar
            Application.StatusBar = vizName & " generated: " & Int(10000 * (i / totalRows)) / 100 & "%"
            DoEvents
        End If
        
        Height = MaxBarHeight * bucketValue / maxValue + 0.001
        If Not categorizations Then 'Create a style for this thematic map colouring
            fastConcat KMLData, vbNewLine & "<Style id=""R" & row & "C" & col & """><PolyStyle><color>", catOffset
            If bucketValue > 0 Then
                fastConcat KMLData, "E91000" & Right(Format(Hex((255 * Log(1 + (Exp(1) - 1) * bucketValue / maxValue))), "\0@"), 2), catOffset
            Else
                fastConcat KMLData, "66805050", catOffset
            End If
            fastConcat KMLData, "</color><outline>0</outline></PolyStyle></Style>", catOffset
        End If

        If bucketValue > 0 Or Not categorizations Then
            putGrid = True
            fastConcat KMLData, vbNewLine & "<Placemark>" & "<name>" & row & ":" & col & _
                               IIf(categorizations, ":" & bar, "") & "</name>", catOffset
            If categorizations Then 'Refer to the colour selected on the tools sheet
                fastConcat KMLData, "<styleUrl>" & color(bar) & "</styleUrl>", catOffset
                tempDescStr = tempDescStr & "<b>" & barType(bar) & ":</b> " _
                                          & Format(bucketValue, formatStr) _
                                          & " in " & bucket.Count & " location(s)<br/>"
            Else 'Refer to the colour style just generated for this item
                fastConcat KMLData, "<styleUrl>#R" & row & "C" & col & "</styleUrl>" _
                                  & "<description><![CDATA[Contains " & bucket.Count & " location(s)<br/>" & IIf(aggType = 1, "Total ", "Average ") _
                                  & valueName & ": " & Format(bucketValue, formatStr) & "]]></description>", catOffset
            End If
            fastConcat KMLData, "<Polygon><extrude>1</extrude><altitudeMode>" & altitudeMode & "</altitudeMode>" _
                              & "<outerBoundaryIs><LinearRing>", catOffset
            x1 = minLong + unit * (col + bar / numbars)
            x2 = x1 + unit / numbars
            y1 = minLat + unit * (row + bar / numbars)
            y2 = y1 + unit / numbars
            fastConcat KMLData, coordsString(x1, y1, x2, y2, Height) _
                              & "</LinearRing></outerBoundaryIs></Polygon></Placemark>", catOffset
        End If

        If categorizations And putGrid And currBar = barsHere Then 'Draw the white grid denoting the area of a cell
            fastConcat KMLData, vbNewLine & "<Placemark> <style><LineStyle><color>ccffffff</color></LineStyle></style>" _
                              & "<LineString><extrude>0</extrude><tessellate>0</tessellate>" _
                              & "<altitudeMode>clampToGround</altitudeMode>", catOffset
            x1 = minLong + (col * unit)
            x2 = x1 + unit
            y1 = minLat + (row * unit)
            y2 = y1 + unit
            fastConcat KMLData, coordsString(x1, y1, x2, y2, 1) & "</LineString></Placemark> ", catOffset
            fastConcat KMLData, "<Placemark>" _
                                    & "<name>" & row & ":" & col & "</name>" _
                                    & "<description><![CDATA[" & tempDescStr & "]]></description>" _
                                    & "<Style><IconStyle><scale>0.5</scale><Icon>" _
                                    & "<href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href>" _
                                    & "</Icon></IconStyle></Style><Point><coordinates>" _
                                    & (minLong + (col * unit)) & "," & (minLat + (row * unit)) & ",1 " _
                                    & "</coordinates></Point>" _
                              & "</Placemark>", catOffset
            putGrid = False
        End If
        
        currBar = currBar + 1
    Next i
    fastConcat KMLData, "</Document></kml>", catOffset 'End of the kml document
    Application.StatusBar = "Sending data to Google Earth..."
    Automation False
    On Error GoTo tryfile
    GEI.LoadKmlData Trim(KMLData)
    Application.StatusBar = False
    Exit Sub
    
tryfile: 'Try writing data to a temporary file if the other method failed.
    If MsgBox("There is too much data, and Google Earth won't load it directly. Do you want to try to " _
            & "write data to a temporary file and load it? (This error may also occur if Google Earth was closed.", _
            vbYesNo, "Error Communicating with Google Earth") = vbYes Then
        Dim Fnum As Integer
        Dim Fname As String
        Fnum = FreeFile()
        Fname = "C:\Temp\" & vizName & ".kml"
        Open Fname For Output As #Fnum
        Print #Fnum, KMLData
        Close #Fnum
        Shell (CStr("C:\Program Files\Google\Google Earth\googleearth.exe") & " " & Fname)
        MsgBox "Hopefully that worked. The file is at " & Fname & ", delete it when finished", , "Gave it my best shot"
    End If
    Application.StatusBar = False
End Sub

Public Function strColorStyles(ByRef colors() As String) As String
    strColorStyles = vbNewLine
    Dim thisColor As String
    Dim i As Long
    For i = LBound(colors) To UBound(colors)
        thisColor = Mid$(colors(i), 2)
addColour:
        If InStr(1, strColorStyles, thisColor) = 0 Then
            strColorStyles = strColorStyles & "<Style id='" & thisColor & _
                "'><PolyStyle><color>" & thisColor & "</color></PolyStyle></Style>" & vbNewLine
            thisColor = "FF" & Mid$(thisColor, 3) 'If translucent, also add the opaque version of this colour
            GoTo addColour
        End If
    Next i
End Function

'This routine gathers the location data and aggregates it based on location and any categories provided
'It has two variations, one for bucketing data within each region, and one for simply aggregating by region.
Private Sub retrieveLocationsData(ByRef categorized As Boolean, Optional ByVal strArrayBarTypes As Variant)
    Dim minLat As Double, minLon As Double, maxLat As Double, maxLon As Double
    Dim tempLat As Double, tempLong As Double 'Stores the lat/long of a row in the locations sheet
    Dim lastRow As Long, numbars As Integer 'Number of rows and number of categories
    Dim i As Long, j As Integer 'Used for row counter and category counter
    Dim col As Long, row As Long, bar As Integer 'Loop counters for regions and grouping categories
    Dim barType() As String, currType As String 'The categorizations and categorization of the current row
    Dim unit As Double 'The size of location groupings based on the tools control
    Dim bucket As DataBucket
    Dim groupAllIndex As Integer 'In the event of a category that contains all un-grouped exposures
    'The column number of various pieces of data retrieved from the worksheet
    Dim valueCol As Integer: valueCol = getSelectedPlotValueCol()
    Dim lattitudeCol As Integer: lattitudeCol = colLattitude()
    Dim longitudeCol As Integer: longitudeCol = colLongitude()
    Dim categoryCol As Integer: categoryCol = getSelectedPlotCategoryCol()
    Dim colID As Integer: colID = colUniqueID()
    Dim aggType As Integer
    If optSum() = xlOn Then aggType = 1 Else If optAverage() = xlOn Then aggType = 2
    
    If Not IsMissing(strArrayBarTypes) Then barType = strArrayBarTypes
    If IsMissing(strArrayBarTypes) And categorized Then _
        Err.Raise 1, , "Cannot retrieve locations as categorized without providing categories."
    'Get the number of categories per region
    If categorized Then
        numbars = UBound(barType) - LBound(barType) + 1
        'Detect whether a category exists for grouping all remaining exposures
        groupAllIndex = -1
        i = LBound(barType)
        For i = LBound(barType) To UBound(barType)
            If barType(i) = "(All Remaining Locations)" Then groupAllIndex = i
            If LCase(barType(i)) = "(blank)" Then barType(i) = ""
        Next i
    Else
        numbars = 1
    End If
    unit = ToolsControlsRoutines.getUnitValue()
    'Get the co-ordinate boundaries of the available locations
    minLat = Application.WorksheetFunction.Min(toEndOfColumn(wkstLocations.Cells(2, lattitudeCol)))
    maxLat = Application.WorksheetFunction.Max(toEndOfColumn(wkstLocations.Cells(2, lattitudeCol)))
    minLon = Application.WorksheetFunction.Min(toEndOfColumn(wkstLocations.Cells(2, longitudeCol)))
    maxLon = Application.WorksheetFunction.Max(toEndOfColumn(wkstLocations.Cells(2, longitudeCol)))
    'Get information for the loop
    lastRow = endOfColumn(wkstLocations, colID).row
    'Collect the data
    ToolsControlsRoutines.RememberControlSettings
    maxValue = 0
    Application.StatusBar = "Gathering Exposure Data"
    Set Buckets = New Map3D
    If Not categorized Then bar = 0 'Category bucket is always 0 (no category)
    For i = 2 To lastRow
        If lastRow < 1000 Or i Mod 1000 = 1 Then
            Application.StatusBar = "Gathering Exposure Data: " & (i - 1) & "/" & (lastRow - 1) & _
                            "(" & Int(10000 * ((i - 1) / (lastRow - 1))) / 100 & "%) complete..."
            DoEvents
        End If
        If Not IsEmpty(wkstLocations.Cells(i, lattitudeCol).value) Then
            tempLat = wkstLocations.Cells(i, lattitudeCol).value
            If Not IsEmpty(wkstLocations.Cells(i, longitudeCol).value) Then
                tempLong = wkstLocations.Cells(i, longitudeCol).value
                'Lat and long are valid, proceed.
                If categorized Then 'Determine the column bucket
                    currType = wkstLocations.Cells(i, categoryCol)
                    bar = -1
                    For j = LBound(barType) To UBound(barType)
                        If barType(j) = currType Then bar = j - LBound(barType)
                    Next j
                    If bar = -1 And groupAllIndex <> -1 Then bar = groupAllIndex
                End If
                If bar <> -1 Then
                    'Determine the location bucket
                    row = Int((tempLat - minLat) / unit)
                    col = Int((tempLong - minLon) / unit)
                    'Add values to bucket
                    If Buckets.Exists(row, col, bar) Then
                        Set bucket = Buckets.Item(row, col, bar)
                    Else
                        Set bucket = New DataBucket
                        bucket.i = row: bucket.j = col: bucket.k = bar
                        Buckets.Add bucket, row, col, bar
                    End If
                    bucket.Add CDbl(wkstLocations.Cells(i, valueCol).value)
                    If aggType = 1 Then
                        If bucket.Sum > maxValue Then maxValue = bucket.Sum
                    End If
                End If
            End If
        End If
    Next i
    'If values are averages the max can't be found on the fly (because a later addition might
    'decrease the average value of that bucket) so they must be computed at the end
    If aggType = 2 Then
        Application.StatusBar = "Finding Largest Average Value"
        maxValue = 0
        Dim elem As Variant
        For Each elem In Buckets.toArray
            Set bucket = elem
            If bucket.Average > maxValue Then maxValue = bucket.Average
        Next elem
    End If
    Set bucket = Nothing
    'Store values for quick generation of new graphs with the same data
    OptRecycleDataEnabled True
    optRecycleData().value = xlOn
    dataRecalled = True
End Sub

'Creates a string with the coordinates specifying the square defined by (x1,y1) and (x2,y2)
Private Function coordsString(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                              Optional ByVal Height As Double = 1) As String
    coordsString = "<coordinates>" _
                      & x1 & "," & y1 & "," & Height & " " _
                      & x2 & "," & y1 & "," & Height & " " _
                      & x2 & "," & y2 & "," & Height & " " _
                      & x1 & "," & y2 & "," & Height & " " _
                      & x1 & "," & y1 & "," & Height & _
                   "</coordinates>"
End Function

