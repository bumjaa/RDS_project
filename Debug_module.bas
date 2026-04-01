Attribute VB_Name = "debug_module"
Sub hidden_sheet()

    Application.ScreenUpdating = False

    'Sheet3.Visible = xlSheetVeryHidden     'Version
    Sheet24.Visible = xlSheetVeryHidden    'db
    Sheet4.Visible = xlSheetVeryHidden     '9832 9835
    Sheet5.Visible = xlSheetVeryHidden     '9814
    Sheet6.Visible = xlSheetVeryHidden     '9610-6-4 -2
    Sheet7.Visible = xlSheetVeryHidden     '9610-6-3 -1
    Sheet8.Visible = xlSheetVeryHidden     '60255-26
    Sheet9.Visible = xlSheetVeryHidden     '9815 9547
    Sheet10.Visible = xlSheetVeryHidden    '9992
    Sheet11.Visible = xlSheetVeryHidden    '3124
    Sheet12.Visible = xlSheetVeryHidden    '3143
    Sheet13.Visible = xlSheetVeryHidden    '3124 new
    Sheet14.Visible = xlSheetVeryHidden    '8102-1 -2
    Sheet15.Visible = xlSheetVeryHidden    '60947-1
    Sheet16.Visible = xlSheetVeryHidden    '61131-2
    Sheet17.Visible = xlSheetVeryHidden    '50 51
    Sheet18.Visible = xlSheetVeryHidden    '60947-2
    Sheet19.Visible = xlSheetVeryHidden    '3369
    Sheet20.Visible = xlSheetVeryHidden    'EN 32 35
    Sheet21.Visible = xlSheetVeryHidden    'EN 301 489-1
    Sheet22.Visible = xlSheetVeryHidden    'FCC PART 15 B
    Sheet23.Visible = xlSheetVeryHidden    'Marine
    Sheet25.Visible = xlSheetVeryHidden    'EN 14
    Sheet26.Visible = xlSheetVeryHidden    '
    Sheet27.Visible = xlSheetVeryHidden    'EN 14
    Sheet28.Visible = xlSheetVeryHidden    'KS C 9800-3
    Sheet29.Visible = xlSheetVeryHidden    'KS C 60974-4-1
    Sheet30.Visible = xlSheetVeryHidden    'VCCI CISPR 32
    
    
    Application.ScreenUpdating = True

End Sub

Sub unhidden_sheet()

    Dim ws As Worksheet
    
    Application.ScreenUpdating = False

    Sheet1.Visible = xlSheetVisible
    Sheet2.Visible = xlSheetVisible
    'Sheet3.Visible = xlSheetVisible
    Sheet24.Visible = xlSheetVisible

    'Sheet4.Visible = xlSheetVisible     '9832 9835
    'Sheet5.Visible = xlSheetVisible     '9814
    'Sheet6.Visible = xlSheetVisible     '9610-6-4 -2
    'Sheet7.Visible = xlSheetVisible     '9610-6-3 -1
    Sheet8.Visible = xlSheetVisible     '60255-26
    'Sheet9.Visible = xlSheetVisible     '9815 9547
    'Sheet10.Visible = xlSheetVisible    '9992
    'Sheet11.Visible = xlSheetVisible    '3124
    'Sheet12.Visible = xlSheetVisible    '3143
    'Sheet13.Visible = xlSheetVisible    '3124 new
    'Sheet14.Visible = xlSheetVisible    '8102-1 -2
    'Sheet15.Visible = xlSheetVisible    '60947-1
    'Sheet16.Visible = xlSheetVisible    '61131-2
    'Sheet17.Visible = xlSheetVisible    '50 51
    'Sheet18.Visible = xlSheetVisible    '60947-2
    'Sheet19.Visible = xlSheetVisible    '3369
    'Sheet20.Visible = xlSheetVisible    'EN 32 35
    'Sheet21.Visible = xlSheetVisible    'EN 301 489-1
    'Sheet22.Visible = xlSheetVisible    'FCC PART 15 B
    'Sheet23.Visible = xlSheetVisible    'Marine
    'Sheet25.Visible = xlSheetVisible    'EN 14
    'Sheet26.Visible = xlSheetVisible    'EN 14
    'Sheet27.Visible = xlSheetVisible    'EN 14
    'Sheet28.Visible = xlSheetVisible    'KS C 9800-3
    'Sheet29.Visible = xlSheetVisible    'KS C 60974-4-1
    Sheet30.Visible = xlSheetVisible
    
    Application.ScreenUpdating = True

End Sub

Sub shhetInitialization()

    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim keep1 As String, keep2 As String

    
    keep1 = "Main"
    keep2 = "Layout"
    
    Sheet1.Visible = xlSheetVisible
    Sheet2.Visible = xlSheetVisible
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> keep1 And ws.Name <> keep2 Then
            ws.Visible = xlSheetVeryHidden
        Else
            ws.Visible = xlSheetVisible
        End If
    Next ws
    
    Sheet1.Activate

End Sub

