Attribute VB_Name = "Updater"
Option Explicit

Const VersionURL = "https://raw.github.com/alainbryden/Google-Earth-Visualization/master/Version.txt"
Const ChangesURL = "https://raw.github.com/alainbryden/Google-Earth-Visualization/master/VersionChange.txt"
Const LatestVersionURL = "https://github.com/alainbryden/Google-Earth-Visualization"

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Public Sub CheckVersion()
    On Error GoTo fail
    Application.StatusBar = "Checking for newer version..."

    Dim ThisVersion As String, LatestVersion As String, VersionChanges As String
    ThisVersion = Tools.Shapes("CurrentVersion").DrawingObject.Text
    
    LatestVersion = FetchFile(VersionURL)
    VersionChanges = FetchFile(ChangesURL)
    If LatestVersion = vbNullString Then
        Application.StatusBar = "Version Check Failed!"
        Exit Sub
    Else
        If LatestVersion = ThisVersion Then
            Application.StatusBar = "Version Check: You are running the latest version!"
        Else
            Application.StatusBar = "Version Check: This visualization tool is out of date!"
            If (MsgBox("You are not running the latest version of this tool. Your version is " & _
                ThisVersion & ", and the latest version is " & LatestVersion & vbNewLine & _
                vbNewLine & "Changes: " & VersionChanges & vbNewLine & _
                vbNewLine & "Click OK to visit the latest version download link.", vbOKCancel, _
                "Tool Out of Date Notification") = vbOK) Then
                ShellExecute 0, vbNullString, LatestVersionURL, vbNullString, vbNullString, vbNormalFocus
            End If
        End If
    End If
    Exit Sub
fail:
    Application.StatusBar = "Version Check Failed (" & Err.Description & ")"
End Sub

Private Function FetchFile(ByVal URL As String) As String
    Dim oHTTP As WinHttp.WinHttpRequest
    
    FetchFile = vbNullString

    Set oHTTP = New WinHttp.WinHttpRequest
    oHTTP.Open Method:="GET", URL:=URL, async:=False
    oHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    oHTTP.setRequestHeader "Content-Type", "text/plain; "
    oHTTP.Option(WinHttpRequestOption_EnableRedirects) = True
    oHTTP.send

    Dim success As Boolean
    success = oHTTP.waitForResponse()
    If Not success Then Exit Function

    FetchFile = oHTTP.responseText
    Set oHTTP = Nothing
End Function


