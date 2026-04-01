Attribute VB_Name = "standard_module"
Option Explicit

' ============================================================
' standard_module: Standard-to-function/suffix mapping
' - Data loaded from standards_data.json (no more giant Case statements)
' - Alias normalization handles comma/space variants automatically
' ============================================================

Private pStdData As Object       ' Dictionary: normalizedAlias -> index
Private pStdList As Collection   ' Collection of parsed standard objects
Private pLoaded  As Boolean

' ── Load standards data from JSON file ────────────────────────
Private Sub EnsureStandardsLoaded()
    If pLoaded Then Exit Sub

    Dim fso As Object, ts As Object
    Dim jsonPath As String, raw As String
    Dim parsed As Object, stds As Collection
    Dim i As Long, j As Long
    Dim stdObj As Object, aliases As Collection
    Dim normalizedAlias As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    jsonPath = ThisWorkbook.Path & "\standards_data.json"

    If Not fso.FileExists(jsonPath) Then
        pLoaded = True
        Exit Sub
    End If

    Set ts = fso.OpenTextFile(jsonPath, 1, False, -1)
    raw = ts.ReadAll
    ts.Close

    Set parsed = JsonConverter.ParseJson(raw)
    Set stds = parsed("standards")

    Set pStdData = CreateObject("Scripting.Dictionary")
    pStdData.CompareMode = vbTextCompare
    Set pStdList = stds

    For i = 1 To stds.Count
        Set stdObj = stds(i)
        Set aliases = stdObj("aliases")
        For j = 1 To aliases.Count
            normalizedAlias = NormalizeStdName(CStr(aliases(j)))
            If Not pStdData.Exists(normalizedAlias) Then
                pStdData.Add normalizedAlias, i
            End If
        Next j
    Next i

    pLoaded = True
End Sub

' ── Normalize standard name (handles comma/space variants) ────
Private Function NormalizeStdName(ByVal s As String) As String
    s = Trim$(s)
    ' Remove extra spaces
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    ' Normalize comma+space → ", "
    s = Replace(s, " ,", ",")
    s = Replace(s, ", ", ",")
    s = Replace(s, ",", ", ")
    NormalizeStdName = s
End Function

' ── Lookup helper: find standard index by name ────────────────
Private Function FindStdIndex(ByVal std As Variant) As Long
    EnsureStandardsLoaded
    Dim normalized As String
    normalized = NormalizeStdName(CStr(std))

    If pStdData Is Nothing Then
        FindStdIndex = 0
    ElseIf pStdData.Exists(normalized) Then
        FindStdIndex = pStdData(normalized)
    Else
        FindStdIndex = 0
    End If
End Function

' ── Public API: Get test functions for a standard ─────────────
Function standardFunction(std) As Variant
    Dim idx As Long
    idx = FindStdIndex(std)

    If idx = 0 Then
        standardFunction = Array()
        Exit Function
    End If

    Dim stdObj As Object
    Set stdObj = pStdList(idx)

    Dim funcs As Collection
    Set funcs = stdObj("functions")

    Dim result() As String
    Dim i As Long
    If funcs.Count = 0 Then
        standardFunction = Array("")
        Exit Function
    End If

    ReDim result(0 To funcs.Count - 1)
    For i = 1 To funcs.Count
        result(i - 1) = CStr(funcs(i))
    Next i

    standardFunction = result
End Function

' ── Public API: Get data suffixes for a standard ──────────────
Function DataSuffix(std) As Variant
    Dim idx As Long
    idx = FindStdIndex(std)

    If idx = 0 Then
        DataSuffix = Array()
        Exit Function
    End If

    Dim stdObj As Object
    Set stdObj = pStdList(idx)

    Dim suffixes As Collection
    Set suffixes = stdObj("data_suffixes")

    Dim result() As String
    Dim i As Long
    If suffixes.Count = 0 Then
        DataSuffix = Array()
        Exit Function
    End If

    ReDim result(0 To suffixes.Count - 1)
    For i = 1 To suffixes.Count
        result(i - 1) = CStr(suffixes(i))
    Next i

    DataSuffix = result
End Function
