Attribute VB_Name = "version_module"
' ============================================================
' version_module: GitHub based version check
' - Uses Config_module for token/URL, Http_module for requests
' - Simplified: single version source from config.json
' ============================================================

Sub CheckVersionFromGithub()
    Dim latestVersion As String
    Dim currentVersion As String
    Dim targetURL As String
    Dim token As String

    currentVersion = GetVersion()
    targetURL = "https://raw.githubusercontent.com/" & _
                CStr(GetCfg("github.repo")) & "/refs/heads/main/version.txt"

    On Error GoTo ErrorHandler
    latestVersion = Trim(HttpGet(targetURL))

    If latestVersion <> currentVersion Then
        MsgBox "New update available! (Latest: " & latestVersion & ")", vbInformation
    Else
        MsgBox "You are on the latest version.", vbInformation
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Network error or invalid URL.", vbCritical
End Sub


Sub FinalAutoUpdate()
    Dim token As String
    Dim versionURL As String, fileURL As String
    Dim latestVer As String, currentVer As String
    Dim savePath As String, targetSheetName As String
    Dim remoteWb As Workbook, localWb As Workbook
    Dim targetSheet As Worksheet

    token = GetGitHubToken()
    currentVer = GetVersion()

    versionURL = CStr(GetCfg("github.version_url", _
                 "https://api.github.com/repos/" & CStr(GetCfg("github.repo")) & "/contents/version.txt"))
    fileURL = "https://api.github.com/repos/" & CStr(GetCfg("github.repo")) & _
              "/contents/RDS_original_260212_working.xlsm"

    savePath = Environ("USERPROFILE") & "\Downloads\Updated_Result_Form.xlsm"
    targetSheetName = "KS C 9800-3"
    Set localWb = ThisWorkbook

    On Error GoTo ErrorHandler

    ' [Step 1] Version check
    latestVer = Trim(HttpGet(versionURL, token, "application/vnd.github.v3.raw"))

    If latestVer = currentVer Then
        MsgBox "Already on latest version.", vbInformation
        Exit Sub
    End If

    ' [Step 2] Download new file
    If MsgBox("New version (" & latestVer & ") available. Update?", vbYesNo) = vbYes Then
        HttpDownloadFile fileURL, savePath, token

        ' [Step 3] Sheet replacement
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        Set remoteWb = Workbooks.Open(savePath)

        On Error Resume Next
        localWb.Sheets(targetSheetName).Delete
        On Error GoTo 0

        remoteWb.Sheets(targetSheetName).Copy After:=localWb.Sheets(localWb.Sheets.Count)
        Set targetSheet = localWb.Sheets(targetSheetName)

        Dim nm As Name
        For Each nm In localWb.Names
            If InStr(1, nm.RefersTo, remoteWb.Name) > 0 Then
                On Error Resume Next
                nm.RefersTo = Replace(nm.RefersTo, "[" & remoteWb.Name & "]", "")
                On Error GoTo 0
            End If
        Next nm

        Dim linkSources As Variant
        linkSources = localWb.LinkSources(xlExcelLinks)
        If Not IsEmpty(linkSources) Then
            Dim i As Integer
            For i = 1 To UBound(linkSources)
                If InStr(1, linkSources(i), remoteWb.Name) > 0 Then
                    localWb.BreakLink Name:=linkSources(i), Type:=xlLinkTypeExcelLinks
                End If
            Next i
        End If

        remoteWb.Close SaveChanges:=False
        Kill savePath

        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        MsgBox "'" & targetSheetName & "' updated to version " & latestVer & "!", vbInformation

    End If

    Exit Sub

ErrorHandler:
    If Not remoteWb Is Nothing Then
        On Error Resume Next
        remoteWb.Close SaveChanges:=False
    End If

    MsgBox "Error: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
