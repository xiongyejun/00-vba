VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'作者QQ 648555205
'淘宝店铺 https://item.taobao.com/item.htm?spm=686.1000925.0.0.H2w0Qj&id=534896704968

Private StrSourceFile As String
Private strSql As String
Private IfErr As Boolean

'Microsoft ActiveX Data Objects 2.8 Library
Dim AdoConn As Object ' ADODB.Connection
Dim rst As Object

'    Dim c_ado As CADO
'    Dim str_sql As String
'
'    Set c_ado = New CADO
'    c_ado.SourceFile = ThisWorkbook.FullName
'    c_ado.SQL = str_sql
'    c_ado.ExcuteSql
'
'    Set c_ado = Nothing

Property Let SQL(str As String)
    strSql = str
End Property

Property Let SourceFile(FileName As String)
    StrSourceFile = FileName
   
    AdoOpen
End Property

Private Function Provider() As String
'A Left Join [Excel 12.0;Database=" & ThisWorkbook.FullName & ";].[" & ActiveSheet.Name & "$] B On

    If VBA.Right$(StrSourceFile, 4) = ".xls" Then
        Provider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrSourceFile
        Provider = Provider & ";Extended Properties=""Excel 8.0;HDR=YES"";"
       
    ElseIf VBA.Right$(StrSourceFile, 4) = ".mdb" Then
        Provider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrSourceFile
       
    ElseIf VBA.Right$(StrSourceFile, 5) = ".xlsx" Or VBA.Right$(StrSourceFile, 5) = ".xlsm" Then
        Provider = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & StrSourceFile
        Provider = Provider & ";Extended Properties=""Excel 12.0;HDR=YES"";"
       
    ElseIf VBA.Right$(StrSourceFile, 6) = ".accdb" Then
        Provider = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & StrSourceFile
       
    ElseIf VBA.Right$(StrSourceFile, 4) = ".txt" Then
        'StrSourceFile应该为文件的路径
        StrSourceFile = VBA.Left(StrSourceFile, VBA.InStrRev(StrSourceFile, "\"))
       
        Provider = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & StrSourceFile
        Provider = Provider & ";Extended Properties=""TEXT;HDR=YES"";"

    'ElseIf VBA.Right$(StrSourceFile, 4) = ".txt" Then
    '    Provider = "Provider=SQLOLEDB;Initial Catalog=" & myDataBase & ";User ID=" & myname & ";Password=" & mypassword & ";Data Source=" & serverip
    'serverip"127.0.0.1"，表示本机，如果数据库在其他电脑上，就写那个电脑的IP
    'Integrated Security=SSPI;Persist Security Info=True;
    End If
   
End Function

Private Function AdoOpen()
    On Error GoTo errHandle
   
    If AdoConn.State = 1 Then AdoConn.Close 'adStateOpen
    AdoConn.Open Provider()
   
    Exit Function
errHandle:
    MsgBox Err.Description
    IfErr = True
End Function

Private Function RstOpen(Optional iCursorType As Long = 0)
    If rst.State = 1 Then rst.Close
'   adOpenUnspecified=-1 不作特别指定
'   adOpenForwardOnly=0  默认值，前滚静态光标。这种光标只能向前浏览记录集，比如用MoveNext向前滚动，
'                         这种方式可节省资源，提高浏览速度，但诸如BookMark、RecordCount、AbsolutePosition
'                         AbsolutePage都不能使用?

'   adOpenKeyset=1        键集游标，采用这种光标的记录集看不到其它用户的新增、删除操作，
'                         但对于更新原有记录的操作对你是可见的

'   adOpenDynamic=2     动态光标，所有数据库的操作都会立即在用户记录集上反应出来。

'   adOpenStatic=3         静态游标。它为记录产生一个静态备份，其他用户的新增、删除、更新操作对你的记录集来说是不可见的。
    rst.Open strSql, AdoConn, iCursorType, 1 ' adOpenForwardOnly  0, adLockReadOnly
End Function

Function ExcuteSQL() As String
    On Error GoTo errHandle
   
    If Not IfErr Then AdoConn.Execute strSql, , 1 Or 128  'adCmdText Or adExecuteNoRecords
    Exit Function
   
errHandle:
    ExcuteSQL = Err.Description
    IfErr = True
End Function

Function ResultToArr(arr() As Variant) As String
    Dim i As Long
   
    On Error GoTo Err
   
    If Not IfErr Then
        RstOpen
        arr = rst.GetRows()
    End If

    Exit Function
Err:
    ResultToArr = Err.Description
    IfErr = True
End Function

Function ResultToExcel(Rng As Range, Optional IfNeedField As Boolean = True) As String
    Dim i As Long
   
    On Error GoTo Err
   
    If Not IfErr Then
        RstOpen
       
        If IfNeedField Then
            For i = 0 To rst.Fields.Count - 1
                Rng.Offset(0, i).Value = rst.Fields(i).Name
            Next i
            Set Rng = Rng.Offset(1, 0)
        End If
       
        Rng.CopyFromRecordset rst 'AdoConn.Execute(SqlStr)
    End If
   
    Exit Function
   
Err:
    ResultToExcel = Err.Description
    IfErr = True
End Function

'str_sql        sql语句
'source_name    数据源名称
'wk             需要添加透视表的工作簿
'rng            透视表的位置
'pc_table_name  透视表的名称
Function CreatePivotCache(wk As Workbook, Rng As Range, pc_table_name As String) As String
    Dim pc As PivotCache
       
    On Error GoTo Err
   
    If Not IfErr Then
        RstOpen
        
        Set pc = wk.PivotCaches.Add(xlExternal)
        Set pc.Recordset = rst
        pc.CreatePivotTable Rng, pc_table_name
        
    End If
   
    Exit Function
   
Err:
    CreatePivotCache = Err.Description
    IfErr = True
End Function

Function GetFieldInfo(fieldNameArr() As String, field_type() As Long) As Long '-1表示出错，k表示字段数量
    Dim i As Long
    Dim k As Long
   
    On Error GoTo Err
   
    If Not IfErr Then
        RstOpen
        k = rst.Fields.Count
       
        ReDim fieldNameArr(k - 1) As String
        ReDim field_type(k - 1) As Long
        For i = 0 To k - 1
            field_type(i) = rst.Fields(i).Type '202 文本  7 日期
            fieldNameArr(i) = rst.Fields(i).Name
        Next i
       
        GetFieldInfo = k
    End If
   
    Exit Function
   
Err:
    MsgBox Err.Description
    GetFieldInfo = -1
    IfErr = True
End Function

'检查字段strField是否存在Value的值
Function ExistsValue(TableName As String, strField As String, Value As Variant) As Boolean
    On Error GoTo Err
    Dim str As String
    
    str = "Select Top 1 " & strField & " From " & TableName & " Where " & strField & "="
    If Not IfErr Then
        If VBA.TypeName(Value) = "String" Then
            Me.SQL = str & "'" & VBA.CStr(Value) & "'"
        ElseIf VBA.TypeName(Value) = "Date" Then
            Me.SQL = str & "#" & VBA.Format(Value, "yyyy-mm-dd") & "#"
        Else
            Me.SQL = str & "" & VBA.CStr(Value) & ""
        End If
        RstOpen 1   'adOpenKeyset下才能查看RecordCount的数量

        If rst.RecordCount > 0 Then
            ExistsValue = True
        Else
            ExistsValue = False
        End If
        
    Else
        ExistsValue = False
    End If
   
    Exit Function
   
Err:
    MsgBox Err.Description
    ExistsValue = False
    IfErr = True
End Function

Private Sub Class_Initialize()
    Set AdoConn = CreateObject("ADODB.Connection")
    Set rst = CreateObject("ADODB.Recordset")
    IfErr = False
End Sub

Private Sub Class_Terminate()
    If rst.State = 1 Then rst.Close
    If AdoConn.State = 1 Then AdoConn.Close 'adStateOpen
   
    Set rst = Nothing
    Set AdoConn = Nothing
End Sub

