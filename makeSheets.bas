Attribute VB_Name = "makeSheets"

Sub CopyAndRenameSheet()

    Dim wsSource As Worksheet
    Dim wsCopied As Worksheet
    Dim sheetNameToCopy As String
    Dim newSheetName As String
    Dim stdName As String
    Dim ordNo As String
    Dim ins_val As String
    Dim countCopied As Long
    Dim i As Long
    Dim Preset_val As String

    
    countCopied = 0
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    For i = 1 To 9
        With ThisWorkbook.Worksheets("Main")
            .Activate
            stdName = Range("STD_" & Format(i, "00")).value
            ordNo = Range("OrderNo_" & Format(i, "00")).value
            ins_val = Range("INS_" & Format(i, "00")).value
            Preset_val = Range("PRESET_" & Format(i, "00")).value
            make_val = Range("Make_" & Format(i, "00")).value
            
            ' OrderNo 값이 없거나 "DTNC"로 시작하지 않으면 복사하지 않음
            If ordNo = "" Or Left(ordNo, 4) <> "DTNC" Then
                ' 아무 작업도 하지 않음
            Else
                newSheetName = ordNo & "_" & .Range("No_" & Format(i, "00")).value
                sheetNameToCopy = stdName
                
                On Error Resume Next
                Set wsSource = ThisWorkbook.Worksheets(sheetNameToCopy)
                On Error GoTo 0
                
                If wsSource Is Nothing Then
                    MsgBox "적용되지 않은 규격입니다.", vbExclamation
                    Exit Sub
                End If
                
                If make_val = True Then
                
                    wsSource.Visible = xlSheetVisible
                    wsSource.Copy Before:=ThisWorkbook.Sheets(1)
                    Set wsCopied = ThisWorkbook.Sheets(1)
    
                    ' 복사된 시트 이름을 변경
                    On Error Resume Next
                    wsCopied.Name = newSheetName
                    wsCopied.Range("Order_No").value = ordNo
                    If Err.Number <> 0 Then
                        MsgBox "새 시트 이름 '" & newSheetName & "'이(가) 유효하지 않거나 이미 존재합니다.", vbExclamation
                        On Error GoTo 0
                        Exit Sub
                    End If
                    wsSource.Visible = xlSheetVeryHidden
    
                    Call ProcessPresetFromDB(newSheetName, stdName, Preset_val, ins_val)
                    
                    On Error GoTo 0
                    
                    countCopied = countCopied + 1
                End If
            End If
        End With
    Next i
    
    If countCopied > 0 Then
        MsgBox "작업이 완료되었습니다! 총 " & countCopied & " 개의 시트가 복사되었습니다.", vbInformation
        'Sheet1.Visible = xlSheetVeryHidden
    Else
        MsgBox "규격 및 접수번호 확인", vbExclamation
    End If
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub



