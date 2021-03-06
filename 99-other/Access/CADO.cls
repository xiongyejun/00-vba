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
Private StrSql As String
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
    StrSql = str
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

Private Function RstOpen()
    If rst.State = 1 Then rst.Close
    rst.Open StrSql, AdoConn, 0, 1 ' adOpenForwardOnly, adLockReadOnly
End Function

Sub ExcuteSql()
    On Error GoTo errHandle
   
    If Not IfErr Then AdoConn.Execute StrSql, , 1 Or 128  'adCmdText Or adExecuteNoRecords
    Exit Sub
   
errHandle:
    MsgBox Err.Description
    IfErr = True
End Sub

Sub ResultToArr(Arr())
    Dim i As Long
   
    On Error GoTo Err
   
    If Not IfErr Then
        RstOpen
        Arr = rst.GetRows()
    End If

    Exit Sub
Err:
    MsgBox Err.Description
    IfErr = True
End Sub

Sub ResultToExcel(rng As Range, Optional IfNeedField As Boolean = True)
    Dim i As Long
   
    On Error GoTo Err
   
    If Not IfErr Then
        RstOpen
       
        If IfNeedField Then
            For i = 0 To rst.Fields.Count - 1
                rng.Offset(0, i).Value = rst.Fields(i).Name
            Next i
            Set rng = rng.Offset(1, 0)
        End If
       
        rng.CopyFromRecordset rst 'AdoConn.Execute(SqlStr)
    End If
   
    Exit Sub
   
Err:
    MsgBox Err.Description
    IfErr = True
End Sub

'str_sql        sql语句
'source_name    数据源名称
'wk             需要添加透视表的工作簿
'rng            透视表的位置
'pc_table_name  透视表的名称
Function CreatePivotCache(wk As Workbook, rng As Range, pc_table_name As String) As Long
    Dim pc As PivotCache
       
    On Error GoTo Err
   
    If Not IfErr Then
        RstOpen
        
        Set pc = wk.PivotCaches.Add(xlExternal)
        Set pc.Recordset = rst
        pc.CreatePivotTable rng, pc_table_name
        
    End If
   
    Exit Function
   
Err:
    MsgBox Err.Description
    IfErr = True
End Function

Function GetFieldInfo(fieldNameArr() As String, field_type() As String) As Long '-1表示出错，k表示字段数量
    Dim i As Long
    Dim k As Long
   
    On Error GoTo Err
   
    If Not IfErr Then
        RstOpen
        k = rst.Fields.Count
       
        ReDim fieldNameArr(k - 1) As String
        ReDim field_type(k - 1) As String
        For i = 0 To k - 1
            If rst.Fields(i).Type = 202 Then '文本
                field_type(i) = "'"
            ElseIf rst.Fields(i).Type = 7 Then '日期
                field_type(i) = "#"
            End If
            fieldNameArr(i) = rst.Fields(i).Name
        Next i
       
        GetFieldInfo = k
    End If
   
    Exit Function
   
Err:
    MsgBox Err.Description
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

