Attribute VB_Name = "MRibbon"
Option Explicit

Public rib As IRibbonUI
Public DB_Info As DBInfo

Sub RibbonUI_onLoad(Ribbon As IRibbonUI)
    Set rib = Ribbon
    
    rib.ActivateTab "Access"
    Set DB_Info.Tables = New CArr
End Sub

Sub rb_GetLabel(control As IRibbonControl, ByRef returnedVal)
    If VBA.Len(DB_Info.Path) Then
        returnedVal = DB_Info.Path
    End If
End Sub

'新建access文件
Sub rbNewAccess(control As IRibbonControl)
    Dim FileName As String
    
    On Error GoTo err_handle
    FileName = Application.GetSaveAsFilename(ThisWorkbook.Path & "\dbName.accdb")
    If FileName = "" Then Exit Sub
    If FileName = "False" Then Exit Sub
    
    NewAccess FileName
    SetDBPath FileName
    
    Exit Sub
err_handle:
    MsgBox Err.Description
End Sub
'获取表结构
Sub rbGetTableStruct(control As IRibbonControl)
    shtIO.Activate
    Cells.Clear
    
    Dim c_ado As CADO

    If VBA.Len(DB_Info.table) Then
        Set c_ado = New CADO
        c_ado.SourceFile = DB_Info.Path
        c_ado.SQL = "Select * From " & DB_Info.table & " Where 1=2"
        c_ado.ResultToExcel Range("A1"), True
    
        Set c_ado = Nothing
    End If
End Sub
'添加数据
Sub rbInsertData(control As IRibbonControl)
    shtIO.Activate
    
    Dim c_ado As CADO

    If VBA.Len(DB_Info.table) Then
        Set c_ado = New CADO
        c_ado.SourceFile = DB_Info.Path
        c_ado.SQL = "Insert Into " & DB_Info.table & " Select * From [Excel 12.0;Database=" & ThisWorkbook.FullName & ";].[" & shtIO.Name & "$]"
        c_ado.ExcuteSql
    
        Set c_ado = Nothing
    End If
End Sub

'选择数据库
Sub rbSelectDB(control As IRibbonControl)
    Dim DBPath As String
    
    DBPath = MFunc.GetFileName("选择数据库文件。", "Access|*.accdb;*.mdb")
    
    SetDBPath DBPath
End Sub

Sub rbddTable_getItemCount(control As IRibbonControl, ByRef returnedVal)
    returnedVal = DB_Info.Tables.Count
End Sub

'rxddSelectSheet getItemLabel回调
Sub rbddTable_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    returnedVal = DB_Info.Tables.Item(VBA.CLng(index))
End Sub

Sub rbddTable_getItemId(control As IRibbonControl, index As Integer, ByRef id)
    id = "rbddTable" & index + 1
End Sub

' onAction回调
Sub rbddTable_click(control As IRibbonControl, id As String, index As Integer)
    Call rbddTable_getItemLabel(control, index, DB_Info.table)
End Sub


'新建table
Sub rbAddTable(control As IRibbonControl)
    Dim c_ado As CADO
    Dim str_sql As String
    
    shtNewTable.Activate
    DB_Info.table = Range("B1").Value
    
    If Not FileExists(VBA.CStr(DB_Info.Path), True) Then MsgBox "请选择DB": Exit Sub
    If DB_Info.table = "" Then MsgBox "请在B1输入table名": Exit Sub
    
    Dim i_row As Long, i As Long
    ActiveSheet.AutoFilterMode = False
    i_row = Range("A" & Cells.Rows.Count).End(xlUp).Row
    If i_row < NewTable.RowStart Then MsgBox "没有数据": Exit Sub
    
    Dim Arr As New CArr
    Dim PrimaryKey As New CArr
    For i = NewTable.RowStart To i_row
        Arr.Add VBA.CStr(Cells(i, 1).Value) & " " & VBA.CStr(Cells(i, 2).Value)
        If VBA.CStr(Cells(i, 3).Value) = "主键" Then PrimaryKey.Add VBA.CStr(Cells(i, 1).Value)
    Next i
    
    str_sql = "Create Table " & DB_Info.table & " (" & Arr.Join(",") & ")"
    Set c_ado = New CADO
    c_ado.SourceFile = DB_Info.Path
    c_ado.SQL = str_sql
    c_ado.ExcuteSql

    Set c_ado = Nothing
    
    If PrimaryKey.Count Then SetPrimaryKey PrimaryKey.Items
End Sub

'设置主键
Function SetPrimaryKey(StrField() As String)
    'Microsoft ADO Ext ...
    Dim cat As Object ' New ADOX.Catalog
    Dim table As Object 'New ADOX.table
    
    Set cat = VBA.CreateObject("ADOX.Catalog")
    Set table = VBA.CreateObject("ADOX.Table")
    
    cat.ActiveConnection = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & DB_Info.Path
    Set table = cat.Tables(DB_Info.table)

'    table.Name = DB_Info.table
'    table.Columns.Append "ID", adInteger, 20
'    table.Columns.Append "TextField", adVarWChar, 20
'
    Dim index As Object ' New ADOX.index
    Set index = VBA.CreateObject("ADOX.index")
    index.Name = "PrimaryKey"
    
    Dim i As Long
    For i = 0 To UBound(StrField)
        index.Columns.Append StrField(i)
    Next
    
    index.PrimaryKey = True
    index.Unique = True
    table.Indexes.Append index
'
'    table.Indexes.Append "TextIndex", "TextField"'
'    cat.Tables.Append table
End Function
