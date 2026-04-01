Attribute VB_Name = "Config_module"
Option Explicit

' ============================================================
' Config_module: 외부 설정 파일(config.json) 기반 설정 관리
' - ThisWorkbook.Path\config.json 에서 설정을 읽어 Dictionary에 캐싱
' - GetCfg("api.key") 같은 점 표기법으로 중첩 값 접근
' ============================================================

Private pConfig As Object  ' Scripting.Dictionary (최상위)
Private pLoaded As Boolean

' ── 설정 로드 ──────────────────────────────────────────────
Public Sub LoadConfig()
    Dim fso As Object, ts As Object
    Dim jsonPath As String
    Dim raw As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    jsonPath = ThisWorkbook.Path & "\config.json"

    If Not fso.FileExists(jsonPath) Then
        ' config.json 없으면 config.sample.json 복사 시도
        Dim samplePath As String
        samplePath = ThisWorkbook.Path & "\config.sample.json"
        If fso.FileExists(samplePath) Then
            fso.CopyFile samplePath, jsonPath
            MsgBox "config.json 파일이 생성되었습니다." & vbCrLf & _
                   "config.json을 열어 실제 값으로 수정해 주세요.", vbExclamation
        Else
            MsgBox "config.json 파일을 찾을 수 없습니다." & vbCrLf & _
                   "경로: " & jsonPath, vbCritical
            pLoaded = False
            Exit Sub
        End If
    End If

    Set ts = fso.OpenTextFile(jsonPath, 1, False, -1) ' TristateTrueIsUnicode=-1
    raw = ts.ReadAll
    ts.Close

    Set pConfig = JsonConverter.ParseJson(raw)
    pLoaded = True
End Sub

' ── 설정값 조회 (점 표기법) ──────────────────────────────────
' 예: GetCfg("api.key"), GetCfg("paths.personal_db"), GetCfg("version")
Public Function GetCfg(ByVal dotPath As String, Optional ByVal defaultValue As Variant = "") As Variant
    If Not pLoaded Then LoadConfig
    If Not pLoaded Then
        GetCfg = defaultValue
        Exit Function
    End If

    Dim parts() As String
    Dim node As Object
    Dim i As Long

    parts = Split(dotPath, ".")
    Set node = pConfig

    On Error GoTo Fallback
    For i = LBound(parts) To UBound(parts) - 1
        Set node = node(parts(i))
    Next i

    ' 마지막 키의 값 반환
    Dim lastKey As String
    lastKey = parts(UBound(parts))

    If IsObject(node(lastKey)) Then
        Set GetCfg = node(lastKey)
    Else
        GetCfg = node(lastKey)
    End If
    Exit Function

Fallback:
    GetCfg = defaultValue
End Function

' ── 버전 문자열 반환 ──────────────────────────────────────────
Public Function GetVersion() As String
    GetVersion = CStr(GetCfg("version", "0.0.0"))
End Function

' ── DB 연결 문자열 반환 ───────────────────────────────────────
Public Function GetPersonalDBConn() As String
    Dim dbPath As String
    dbPath = CStr(GetCfg("paths.personal_db", ""))
    If dbPath = "" Then
        Err.Raise vbObjectError + 1, "Config_module", "paths.personal_db가 config.json에 설정되지 않았습니다."
    End If
    GetPersonalDBConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                        "Data Source=" & dbPath & ";" & _
                        "Persist Security Info=False;"
End Function

' ── GitHub 토큰 반환 ──────────────────────────────────────────
Public Function GetGitHubToken() As String
    GetGitHubToken = CStr(GetCfg("github_token", ""))
End Function

' ── API 키 반환 ───────────────────────────────────────────────
Public Function GetApiKey() As String
    GetApiKey = CStr(GetCfg("api.key", ""))
End Function

' ── API URL 반환 ──────────────────────────────────────────────
Public Function GetApiUrl(Optional ByVal urlKey As String = "base_url") As String
    GetApiUrl = CStr(GetCfg("api." & urlKey, ""))
End Function

' ── 설정 캐시 초기화 (설정 파일 변경 후 다시 읽고 싶을 때) ────
Public Sub ReloadConfig()
    pLoaded = False
    Set pConfig = Nothing
    LoadConfig
End Sub

' ── 설정 로드 여부 ────────────────────────────────────────────
Public Function IsConfigLoaded() As Boolean
    IsConfigLoaded = pLoaded
End Function
