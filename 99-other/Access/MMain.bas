Attribute VB_Name = "MMain"
Option Explicit

Enum NewTable
    RowStart = 3
End Enum

Type DBInfo
    Path As String
    table As String
    Tables As CArr
End Type

Sub vba_main()

    
    On Error GoTo err_handle
    
    
    Exit Sub
err_handle:
    MsgBox Err.Description
End Sub


Function NewAccess(FileName As String)
    Dim app As Object ' Access.Application
    
    Set app = VBA.CreateObject("Access.Application")  ' New Access.Application
'    acNewDatabaseFormatAccess12 12 以 Microsoft Access 2010 (.accdb) 文件格式创建数据库。
'    acNewDatabaseFormatAccess2000 9 以 Microsoft Access 2000 (.mdb) 文件格式创建数据库。
'    acNewDatabaseFormatAccess2002 10 以 Microsoft Access 2002-2003 (.mdb) 文件格式创建数据库。
'    acNewDatabaseFormatUserDefault 0 以默认的文件格式创建数据库
    app.NewCurrentDatabase filepath:=FileName, FileFormat:=12 'acNewDatabaseFormatAccess2007
    
    Set app = Nothing
End Function

Sub AddTable(database As String, strTable As String, StrField As String)
    Dim c_ado As CADO
    Dim str_sql As String

    str_sql = "Create Table " & strTable & " (" & StrField & ")"
    Set c_ado = New CADO
    c_ado.SourceFile = database
    
    c_ado.SQL = str_sql
    c_ado.ExcuteSql

    Set c_ado = Nothing
End Sub

Function SetDBPath(DBPath As String)
    If VBA.Len(DBPath) Then
        DB_Info.Path = DBPath
    End If
    
    rib.InvalidateControl "lbDBPath"
    
    Set DB_Info.Tables = New CArr
    '获取DB的tables
    Dim cat As Object ' New ADOX.Catalog
    Dim table As Object 'New ADOX.table
    
    Set cat = VBA.CreateObject("ADOX.Catalog")
    
    cat.ActiveConnection = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & DB_Info.Path
    Dim i As Long
    For i = 0 To cat.Tables.Count - 1
        
        If cat.Tables(i).Type = "TABLE" Then
            DB_Info.Tables.Add cat.Tables(i).Name
        End If
    Next
    
    rib.InvalidateControl "ddTable"
    
    Set cat = Nothing
End Function
