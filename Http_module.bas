Attribute VB_Name = "Http_module"
Option Explicit

' ============================================================
' Http_module: HTTP 요청 헬퍼
' - 반복되는 MSXML2 패턴을 통합
' ============================================================

' ── GET 요청 (텍스트 응답) ─────────────────────────────────────
Public Function HttpGet(ByVal url As String, _
                        Optional ByVal authToken As String = "", _
                        Optional ByVal accept As String = "application/json") As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.Open "GET", url, False
    If accept <> "" Then http.setRequestHeader "Accept", accept
    If authToken <> "" Then http.setRequestHeader "Authorization", "token " & authToken
    http.setRequestHeader "Accept-Encoding", "identity"
    http.send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + http.Status, "Http_module", _
                  "HTTP " & http.Status & ": " & Left$(http.responseText, 300)
    End If

    HttpGet = http.responseText
End Function

' ── GET 요청 (바이너리 응답 → 파일 저장) ─────────────────────
Public Sub HttpDownloadFile(ByVal url As String, _
                            ByVal savePath As String, _
                            Optional ByVal authToken As String = "")
    Dim http As Object, stream As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.Open "GET", url, False
    If authToken <> "" Then
        http.setRequestHeader "Authorization", "token " & authToken
        http.setRequestHeader "Accept", "application/vnd.github.v3.raw"
    End If
    http.send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + http.Status, "Http_module", _
                  "Download failed: HTTP " & http.Status
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open
    stream.Write http.responseBody
    stream.SaveToFile savePath, 2
    stream.Close
End Sub

' ── GET 요청 + JSON 파싱 ──────────────────────────────────────
Public Function HttpGetJson(ByVal url As String, _
                            Optional ByVal authToken As String = "") As Object
    Dim raw As String
    raw = HttpGet(url, authToken, "application/json")
    Set HttpGetJson = JsonConverter.ParseJson(raw)
End Function

' ── Google Apps Script API 호출 헬퍼 ─────────────────────────
' apiUrlKey: config.json의 api 섹션 키 (base_url, instruments_url, usage_url)
' params: "sheet=Test_Preset&key=xxx" 등 쿼리 파라미터
Public Function CallGasApi(ByVal apiUrlKey As String, _
                           ByVal params As String) As String
    Dim baseUrl As String
    baseUrl = GetApiUrl(apiUrlKey)
    If baseUrl = "" Then
        Err.Raise vbObjectError + 1, "Http_module", _
                  "API URL not configured: api." & apiUrlKey
    End If
    CallGasApi = HttpGet(baseUrl & "?" & params)
End Function
