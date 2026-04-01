Attribute VB_Name = "DataInsert_module"
Option Explicit

Sub insertData(ByVal prefix As String, ByVal DataCols As Variant)
    Dim ws               As Worksheet
    Dim rg               As Range, dataRG As Range
    Dim arr              As Variant
    Dim r                As Long, i As Long, c As Long
    Dim isTest           As Boolean, isRF As Boolean
    Dim precom           As String
    Dim commentRangeName As String
    Dim naString         As String
    Dim naComment        As String
    Dim aComment         As String
    
    Application.EnableEvents = False
    On Error GoTo Cleanup
    
    Set ws = ActiveSheet
    
    Select Case Left(ws.Range("STD").value, 2)
        Case "KS", "KN"
            naString = "해당무"
            naComment = "- 해당사항 없음."
            aComment = "- TEST 중 오동작 없이 동작상태를 유지함."
        Case "EN"
            naString = "-"
            naComment = ""
            aComment = "No degradation of performance"
    End Select
    
    isRF = (Left(ws.Range("STD").value, 9) = "KS X 3124") Or (Left(ws.Range("STD").value, 12) = "EN 301 489-1")
    isTest = (ws.Range(prefix).value = 1 Or ws.Range(prefix).value = True)
    
    Set rg = ws.Range(prefix & "_RESULT")
    Set dataRG = rg.Offset(1, 0).Resize(rg.rows.Count - 1, rg.Columns.Count)
    
    arr = dataRG.value
    
    For r = 1 To UBound(arr, 1)
        If Len(Trim(CStr(arr(r, 1)))) > 0 Then
            For i = LBound(DataCols) To UBound(DataCols)
                c = DataCols(i)
                If isRF Then
                    precom = Left(CStr(arr(r, 6)), 5)
                    If isTest Then
                        arr(r, c) = precom & " (A)"
                    Else
                        arr(r, c) = naString
                    End If
                Else
                    If isTest Then
                        arr(r, c) = "A"
                    Else
                        arr(r, c) = naString
                    End If
                End If
            Next i
        End If
    Next r
    
    dataRG.value = arr
    
    commentRangeName = Split(prefix, "_")(0) & "_COMMENTS"
    If isTest Then
        ws.Range(commentRangeName).value = aComment
    Else
        ws.Range(commentRangeName).value = naComment
    End If

Cleanup:
    ' 8) 이벤트 재활성화 및 오류 메시지
    Application.EnableEvents = True
    
End Sub


Public Function FormatHour(val As Variant) As String
    Dim h As Long
    
    If val = "" Then
        Exit Function
    End If
    
    ' 이미 "hh:mm" 형태(예: "05:00")면 그대로 반환 ???
    If VarType(val) = vbString Then
        If val Like "##:##" Then
            FormatHour = val
            Exit Function
        End If
    End If
    
    ' ??? 숫자 판별 ???
    If IsNumeric(val) Then
        h = CLng(val)
        ' 0~24 범위 확인
        If h >= 0 And h <= 24 Then
            FormatHour = Format(h, "00") & ":00"
            Exit Function
        End If
    End If
    
    ' ??? 그 외는 빈 문자열 ???
    FormatHour = ""
    
End Function

