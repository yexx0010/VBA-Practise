Attribute VB_Name = "modStudentMarking"
Global Const DB_Name = "registrar.mdb"
Global Const TB_STUDENT = "data"
Global DB_ConnectString
'Database path and report files path
Global DB_PATH
Global RPT_PATH

'Import csv file to access database through MS ADO
Function ImportTextToAccessADO(txtPath As String, csvFile As String) As Boolean
    Dim cnn As New ADODB.Connection
    Dim sql As String
       
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source= " & DB_PATH & ";" & _
                   "Jet OLEDB:Engine Type=4;"
    
    sql = "INSERT  INTO [data] SELECT * FROM [Text;DATABASE=" & txtPath & "].[" & csvFile & "]"

On Error GoTo importError
    cnn.Execute sql
    ImportTextToAccessADO = True
    Exit Function
    
importError:
    ImportTextToAccessADO = False
    
End Function

'return table record count
Function recCount(chkTable As String) As Long
    Dim i As Integer

    Set conn = New ADODB.Connection
    Set rec = New ADODB.Recordset
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_PATH & ";Persist Security Info=False"
    conn.Open
    
    Set rec = New ADODB.Recordset
    With rec
        .ActiveConnection = conn
        .CursorType = adOpenStatic
        .Open chkTable
    End With
    i = 0
    Do While Not rec.EOF
        i = i + 1
        rec.MoveNext
    Loop
    recCount = i
End Function

'Truncate student table before importing
Sub truncateStudentTable()
    Dim sql As String
    
    Set conn = New ADODB.Connection
    Set rec = New ADODB.Recordset
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_PATH & ";Persist Security Info=False"
    conn.Open
    sql = "DELETE * FROM " & TB_STUDENT
    conn.Execute sql

End Sub
