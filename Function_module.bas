Attribute VB_Name = "Function_module"
Sub ProtectCurrentSheet()
    'ActiveSheet.Protect Password:="dTnC145#"
End Sub

Sub UnprotectCurrentSheet()
    'ActiveSheet.Unprotect Password:="dTnC145#"
End Sub

Sub reRollevent()

    Application.EnableEvents = True
    Application.Volatile True
    
End Sub

Sub stopRollevent()

    Application.EnableEvents = False
    
End Sub


Sub SetEnvMinMaxValue()

    Dim rangeNames As Variant
    Dim suffix As String
    Dim i As Long
    Dim dynamicRange As Range
    Dim currentRangeName As String
    Dim minVal As Date, maxVal As Date
    Dim firstValueFound As Boolean
    Dim dataRange As Range
    Dim cell As Range
    Dim cellValue As Variant
    Dim ws As Worksheet

    Application.EnableEvents = False
    UnprotectCurrentSheet

    Set ws = ActiveSheet

    rangeNames = standardFunction(ws.Range("STD").value)
    suffix = "_ENV"
    firstValueFound = False

    For i = LBound(rangeNames) To UBound(rangeNames)
        currentRangeName = rangeNames(i) & suffix
        On Error Resume Next
        Set dynamicRange = ws.Range(currentRangeName)
        On Error GoTo 0

        If Not dynamicRange Is Nothing Then
            If dynamicRange.rows.Count > 1 Then
                ' 헤더를 제외한 데이터 영역
                Set dataRange = dynamicRange.Offset(1, 0) _
                                .Resize(dynamicRange.rows.Count - 1, dynamicRange.Columns.Count)
                If dataRange.Columns.Count >= 4 Then
                    For Each cell In dataRange.Columns(4).cells
                        cellValue = cell.value
                        If IsDate(cellValue) Then
                            If Not firstValueFound Then
                                minVal = CDate(cellValue)
                                maxVal = CDate(cellValue)
                                firstValueFound = True
                            Else
                                If CDate(cellValue) < minVal Then minVal = CDate(cellValue)
                                If CDate(cellValue) > maxVal Then maxVal = CDate(cellValue)
                            End If
                        End If
                    Next
                End If
            End If
        End If

        ' 다음 루프를 위해 초기화
        Set dynamicRange = Nothing
        Set dataRange = Nothing
    Next i

    If firstValueFound Then
        ws.Range("Test_Period_Start").value = minVal
        ws.Range("Test_Period_End").value = maxVal
    Else
        ws.Range("Test_Period_Start").value = ""
        ws.Range("Test_Period_End").value = ""
    End If
    
    Application.EnableEvents = True
    ProtectCurrentSheet
    
End Sub


Sub AudioLevelApply(ByVal prefix As String)
    Dim ws As Worksheet
    Dim rg As Range
    Dim msg As String
    
    Set ws = ActiveSheet
    ' 접두사를 활용해 SOUND_LEVEL 범위를 지정합니다.
    Set rg = ws.Range(prefix & "_SOUND_LEVEL")
    
    ' 기준(STD)이 맞는 경우에만 실행
    If ws.Range("STD").value = "KS C 9832, KS C 9835" Then
        msg = ""
        
        ' On ear 데이터가 있을 경우 메시지 구성
        If ws.Range(prefix & "_OnEar_L1").value <> "" Then
            msg = "- On ear: L1 - L0 = " & _
                  ws.Range(prefix & "_OnEar_L1").value & " dBm - (" & _
                  ws.Range(prefix & "_OnEar_L0").value & ") dBm = " & _
                  ws.Range(prefix & "_OnEar_Result").value & " dB"
        End If
        
        ' Off ear 데이터가 있을 경우 메시지 추가
        If ws.Range(prefix & "_OffEar_L1").value <> "" Then
            If msg <> "" Then msg = msg & vbLf
            msg = msg & "- Off ear: L1 - L0 = " & _
                  ws.Range(prefix & "_OffEar_L1").value & " dBm - (" & _
                  ws.Range(prefix & "_OffEar_L0").value & ") dBm = " & _
                  ws.Range(prefix & "_OffEar_Result").value & " dB"
        End If
        
        ' on ear, off ear 둘 다 없으면 기본 메시지를 넣음
        If msg = "" Then msg = "- 해당무"
        
        rg.value = msg
    End If
End Sub

