Attribute VB_Name = "Encoder_module"
Option Explicit

' ============================================================
' Encoder_module: URL 인코딩 통합 모듈
' - 기존 Instruments.frm, Peripherals.frm, Ins_module의 중복 제거
' ============================================================

' ── ASCII 전용 URL 인코딩 (간이) ──────────────────────────────
Public Function URLEncode(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9-_.~]" Then
            out = out & ch
        ElseIf ch = " " Then
            out = out & "+"
        Else
            out = out & "%" & Right$("0" & Hex$(Asc(ch)), 2)
        End If
    Next
    URLEncode = out
End Function

' ── UTF-8 URL 인코딩 (한글 등 멀티바이트 지원) ────────────────
' RFC 3986 준수. 기존 Instruments.frm / Peripherals.frm / Ins_module에서
' 각각 복사-붙여넣기되어 있던 함수를 여기로 통합.
Public Function URLEncodeUTF8(ByVal s As String) As String
    Dim i      As Long
    Dim c      As Long
    Dim out    As String
    Dim tmp    As String

    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        Select Case True
            ' ALPHA / DIGIT / unreserved
            Case (c >= &H30 And c <= &H39), _
                 (c >= &H41 And c <= &H5A), _
                 (c >= &H61 And c <= &H7A), _
                 c = &H2D, c = &H2E, c = &H5F, c = &H7E
                out = out & ChrW(c)

            Case c < &H80
                out = out & "%" & Right$("0" & Hex$(c), 2)

            Case c < &H800
                tmp = Hex$((c \ &H40) Or &HC0)
                out = out & "%" & Right$("0" & tmp, 2)
                tmp = Hex$((c And &H3F) Or &H80)
                out = out & "%" & Right$("0" & tmp, 2)

            Case Else
                tmp = Hex$((c \ &H1000) Or &HE0)
                out = out & "%" & Right$("0" & tmp, 2)
                tmp = Hex$(((c \ &H40) And &H3F) Or &H80)
                out = out & "%" & Right$("0" & tmp, 2)
                tmp = Hex$((c And &H3F) Or &H80)
                out = out & "%" & Right$("0" & tmp, 2)
        End Select
    Next

    URLEncodeUTF8 = out
End Function
