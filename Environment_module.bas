Attribute VB_Name = "Environment_module"
Option Explicit

Sub PrintMayMorningData()
    Dim cN     As Object
    Dim rs     As Object
    Dim filePath As String
    Dim connStr As String
    Dim sql     As String

    filePath = CStr(GetCfg("paths.environment_data_dir")) & "\2025\morning.xlsx"

    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & filePath & ";" & _
              "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1;ReadOnly=True"";"

    sql = "SELECT [temp_morning], [humidity_morning], [pressure] " & _
          "FROM [5月$] " & _
          "WHERE [morning] = 1"

    Set cN = CreateObject("ADODB.Connection")
    cN.Open connStr

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, cN, 1, 1

    Do While Not rs.EOF
        Debug.Print "temp: " & rs.Fields("temp_morning").Value & _
                    ", humidity: " & rs.Fields("humidity_morning").Value & _
                    ", pressure: " & rs.Fields("pressure").Value
        rs.MoveNext
    Loop

    rs.Close
    cN.Close
    Set rs = Nothing
    Set cN = Nothing
End Sub
