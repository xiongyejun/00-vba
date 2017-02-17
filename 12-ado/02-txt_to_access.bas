'http://club.excelhome.net/thread-1256518-3-1.html
'如果没有schema.ini，ADO默认文本文件是用逗号分隔的，如果有schema.ini，ADO就按照它指定的分隔符，是否有标题行等进行查询

Function SaveTxtToAccess(sFIELD As String, txt_name As String) As Long
    Dim str_sql As String
    Dim s As String
    Dim mypath
    
    mypath = ThisWorkbook.Path & "\"

    s = "[" & txt_name & "]" & vbCrLf & "COLNAMEHEADER = TRUE" & vbCrLf & "Format = Delimited" & vbCrLf & "Col1=股票代码 Char 6" & vbCrLf & "Col2=股票名称 Char" & _
            vbCrLf & "Col3=日期 DATE" & vbCrLf & "Col4=开盘 Double" & vbCrLf & "Col5=最高 Double" & vbCrLf & "Col6=最低 Double" & vbCrLf & "Col7=收盘 Double" _
             & vbCrLf & "Col8=涨跌幅 Double" & vbCrLf & "Col9=成交量 Double" & vbCrLf & "Col10=成交额 Double"

    Open mypath & "schema.ini" For Output As #1
    Print #1, s
    Close #1
    Kill mypath & "schema.ini"
        
    str_sql = "Insert Into [" & sTABLE_NAME & "] (" & sFIELD & ") " & _
                "Select * From [Text;Database=" & ThisWorkbook.Path & ";].[" & txt_name & "]"

    ExecuteSql str_sql
End Function