Attribute VB_Name = "rename_module"
Sub RenamePDFsBasedOnOperatingMode()
    Dim folderPath As String
    Dim fso As Object, file As Object
    Dim fileName As String, baseName As String
    Dim ws As Worksheet
    Dim opModeRange As Range, i As Long
    Dim modeName As String, modeIndex As String
    Dim prefixList As Variant
    Dim newName As String
    Dim parts As Variant

    Set ws = ActiveSheet
    Set opModeRange = ws.Range("OPERATING_MODE")
    folderPath = ThisWorkbook.path & "\DATA"

    prefixList = Array("CE_", "RE_", "PK_", "AV_")
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        MsgBox "DATA 폴더가 존재하지 않습니다.", vbExclamation
        Exit Sub
    End If

    For Each file In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(file.Name)) = "pdf" Then
            fileName = file.Name
            For Each prefix In prefixList
                If LCase(Left(fileName, Len(prefix))) = LCase(prefix) Then
                    parts = Split(fileName, "_")
                    If UBound(parts) >= 2 Then
                        ' 모드명 파싱 (마지막 부분에서 확장자 제거)
                        modeName = Replace(Split(parts(UBound(parts)), ".")(0), ".pdf", "")
                        
                        ' OPERATING_MODE에서 모드명 찾기
                        For i = 1 To opModeRange.rows.Count
                            If Trim(UCase(opModeRange.cells(i, 2).value)) = Trim(UCase(modeName)) Then
                                modeIndex = opModeRange.cells(i, 1).value ' MODE 1 등
                                ' 새 파일 이름 생성
                                parts(UBound(parts)) = modeIndex & ".pdf"
                                newName = Join(parts, "_")

                                ' 이름이 다르면 변경
                                If newName <> fileName Then
                                    Name folderPath & "\" & fileName As folderPath & "\" & newName
                                    Debug.Print "Renamed: " & fileName & " -> " & newName
                                End If
                                Exit For
                            End If
                        Next i
                    End If
                    Exit For
                End If
            Next prefix
        End If
    Next file

    MsgBox "파일 이름 변경 완료!", vbInformation
End Sub

