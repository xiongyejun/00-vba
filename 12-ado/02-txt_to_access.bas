Attribute VB_Name = "模块1"
Option Explicit

'方便存储不同数据类型
Type StockField
    arr_code_name() As String
    arr_date() As Date
    arr_other() As Double
End Type

Const lDATA_NUM As Long = 1000000
Const sACCESS_NAME As String = "Database1.accdb"
Const sTABLE_NAME As String = "data"

Sub vba_main()
    Dim stock_result As StockField
    Dim t As Double
    
    On Error GoTo err_handle
    
    If DeleteAccess() = 0 Then Exit Sub
    
    RemDimStockResult stock_result
    
    t = GetTxtData(stock_result)
    Application.StatusBar = False
    
    MsgBox "导入数据完成，用时：" & VBA.CLng(Timer - t) & "秒。"
    Exit Sub
err_handle:
    Application.StatusBar = False
    MsgBox Err.Description
End Sub

Function RemDimStockResult(stock_result As StockField)
    ReDim stock_result.arr_code_name(1 To lDATA_NUM, 1 To 2) As String
    ReDim stock_result.arr_date(1 To lDATA_NUM, 1 To 1) As Date
    ReDim stock_result.arr_other(1 To lDATA_NUM, 1 To 7) As Double
End Function

Function GetTxtData(stock_result As StockField) As Double
    Dim str_dir As String
    Dim t As Double
    
    str_dir = GetFolderPath()
    t = Timer
    scan_dir str_dir, stock_result
    
    GetTxtData = t
End Function

Function scan_dir(str_dir As String, stock_result As StockField) As Long
    Dim fso As Object
    Dim file As Object
    Dim folder As Object
    Dim tmp
    Dim k As Long
    Dim k_file As Long
    
    k = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.Getfolder(str_dir)
    
    k_file = 0
    For Each file In folder.Files
        If file.Type = "文本文档" Then
            k_file = k_file + 1
            Application.StatusBar = "正在处理第" & k_file & "个文件。"
            If fso_read_txt(fso, file.Path, stock_result, k) = 0 Then
                Exit Function
            End If
            
            '如果数据行数已经超过了95万，则将数据进行一次保存到access
            If k > 950000 Then
                Application.StatusBar = "正在导入部分数据到access中。"
                
                If SaveDataToAccess(stock_result, k) = 0 Then Exit Function
                RemDimStockResult stock_result
                k = 1
            End If
            
        End If
    Next file
    
    If k > 1 Then
        If SaveDataToAccess(stock_result, k) = 0 Then Exit Function
    End If
    
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Function

Function DeleteAccess() As Long
    Dim str_sql As String
    
    str_sql = "Delete * From [" & sTABLE_NAME & "]"
    
    If ExecuteSql(str_sql) = 0 Then
        DeleteAccess = 0
        Exit Function
    End If
    
    DeleteAccess = 1
End Function

Function providerStr(fileName As String) As String
    providerStr = "Provider=Microsoft.ACE.OLEDB.12.0;" _
        & "Data Source=" & ThisWorkbook.Path & "\" & fileName & ";"
    
End Function

Function ExecuteSql(SqlStr As String) As Long   '0表示出错，1表示正确
    Dim AdoConn As Object
    
    On Error GoTo Err
    Set AdoConn = CreateObject("ADODB.Connection")
    
    AdoConn.Open providerStr(sACCESS_NAME)
    
    AdoConn.Execute (SqlStr)
    
    ExecuteSql = 1
    AdoConn.Close
    Exit Function
    
Err:
    If Err.Number <> 0 Then MsgBox Err.Description
    Set AdoConn = Nothing
    ExecuteSql = 0
End Function


Function SaveDataToAccess(stock_result As StockField, k As Long) As Long
    '先放工作表里，再用sql输入
    Dim arr_col
    
    arr_col = Array("股票代码", "股票名称", "日期", "开盘", "最高", "最低", "收盘", "涨跌幅", "成交量", "成交额")
    Cells.Clear
    Range("A1:J1").Value = arr_col
    Range("A2").Resize(k - 1, 2).Value = stock_result.arr_code_name
    Range("C2").Resize(k - 1, 1).Value = stock_result.arr_date
    Range("D2").Resize(k - 1, 7).Value = stock_result.arr_other
    
    '放入access
    Dim str_sql As String
'    "[Excel 12.0;Database=" & tbSrc_2.Text & ";].[" & tbSrc_2_tabel.Text & "$A:N]"
'    str_sql = "Select * Into [" & sTABLE_NAME & "] From [Excel 12.0;Database=" & ThisWorkbook.FullName & ";].[" & ActiveSheet.Name & "$A:J]"
    str_sql = "Insert Into [" & sTABLE_NAME & "] (" & VBA.Join(arr_col, ",") & ") " & _
                "Select * From [Excel 12.0;Database=" & ThisWorkbook.FullName & ";].[" & ActiveSheet.Name & "$A:J]"
    

    If ExecuteSql(str_sql) = 0 Then
        SaveDataToAccess = 0
        Exit Function
    End If
    Cells.Clear
    
End Function

Function fso_read_txt(fso As Object, file_name As String, stock_result As StockField, k As Long) As Long
    Dim sr As Object
    Dim str As String
    Dim stock_code As String
    Dim stock_name As String
    Dim tmp
    Dim i_col As Long
    Dim start_k As Long, i As Long
    
    On Error GoTo Err
    
    Set sr = fso.OpenTextFile(file_name, 1) 'ForReading=1

    str = sr.ReadLine()
    tmp = VBA.Split(str, " ")
    stock_code = VBA.CStr(tmp(0))
    stock_name = VBA.CStr(tmp(1))
    
    '标题
    str = sr.ReadLine()
    '第1行
    str = sr.ReadLine()
    start_k = k
    Do Until sr.AtEndOfStream
        tmp = VBA.Split(str, vbTab)
        stock_result.arr_code_name(k, 1) = stock_code
        stock_result.arr_code_name(k, 2) = stock_name
        stock_result.arr_date(k, 1) = VBA.CDate(tmp(0))
        
        For i_col = 1 To 4
            stock_result.arr_other(k, i_col) = VBA.CDbl(tmp(i_col))
        Next i_col
        stock_result.arr_other(k, 6) = VBA.CDbl(tmp(5)) '成交量
        stock_result.arr_other(k, 7) = VBA.CDbl(tmp(6)) '成交额
        
        k = k + 1
        str = sr.ReadLine() '最后一行舍弃
    Loop
    
    '涨跌幅
    If stock_result.arr_other(start_k, 1) > 0# Then
        stock_result.arr_other(start_k, 5) = (stock_result.arr_other(start_k, 4) - stock_result.arr_other(start_k, 1)) / stock_result.arr_other(start_k, 1)
    End If
    
    For i = start_k + 1 To k - 1
        If stock_result.arr_other(i, 1) > 0# Then
            stock_result.arr_other(i, 5) = (stock_result.arr_other(i, 4) - stock_result.arr_other(i - 1, 4)) / stock_result.arr_other(i, 1)
        End If
    Next i
    
    fso_read_txt = 1
    Set sr = Nothing
    Exit Function
Err:
    MsgBox "读取文件出错，请检查文件格式：" & vbNewLine & file_name
    fso_read_txt = 0
End Function


Function GetFolderPath() As String
    Dim myFolder As Object
    Set myFolder = CreateObject("Shell.Application").Browseforfolder(0, "选择txt源文件所在文件夹", 0)
    If Not myFolder Is Nothing Then
'        GetFolderPath = myFolder.Items.item.path
        GetFolderPath = myFolder.Self.Path
        If Right(GetFolderPath, 1) <> "\" Then GetFolderPath = GetFolderPath & "\"
    Else
        GetFolderPath = ""
        MsgBox "请选择文件夹。"
    End If
    Set myFolder = Nothing
End Function

