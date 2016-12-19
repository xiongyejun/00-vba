https://www.connectionstrings.com/
[enter link description here](https://www.connectionstrings.com/)

##访问XLSM:##

    Provier=Microsoft Ace OLEDB.12.0;Data Source=带路径的文件名;Extended Properties="Excel 12.0 Macro;HDR=NO";

##访问文本文件(如txt,csv等),数据源只要求指明文件所在的文件夹:##

    Provier=Microsoft Ace OLEDB.12.0;Data Source=文本文件所在文件夹;Extended Properties="text;HDR=YES;FMT=Delimited";


##访问数据Access##
 
  没有密码:
  
    Provier=Microsoft Ace OLEDB.12.0;Data Source=带路径的文件名;
    
  有密码:
   
   Provier=Microsoft Ace OLEDB.12.0;Data Source=带路径的文件名;Jet OLEDB:Database Password=密码;

##MYSQL数据库##

连接到数据库

	Function CnnOpen(ByVal ServerName As String, ByVal DBName As String, ByVal TblName As String, ByVal User As String, ByVal PWD As String) '服务器名或IP、数据库名、登录用户、密码
		Dim CnnStr As String '定义连接字符串
		Set Cnn = CreateObject("ADODB.Connection") '创建ADO连接对象
		
		Cnn.CommandTimeout = 15 '设置超时时间
		CnnStr = "DRIVER={MySql ODBC 5.1 Driver};SERVER=" & ServerName & ";Database=" & DBName & ";Uid=" & User & ";Pwd=" & PWD & ";Stmt=set names GBK" '
		Cnn.ConnectionString = CnnStr
		Cnn.Open
	End Function
	
关闭连接

	Function CnnClose()
		If Cnn.State = 1 Then
			Cnn.Close
		End If
	End Function
		

	Function GetRecordset(ByVal SqlStr As String)
		Set Records = CreateObject("ADODB.recordset")
		Records.CursorType = adOpenStatic '设置游标类型,否则无法获得行数
		Records.CursorLocation = adUseClient '设置游标属性,否则无法获得行数
			
		'对于Connection对象的Execute方法产生的记录集对象，一般是一个只读并且只向前的记录集
		'如果需要对记录集进行操作，譬如修改和增加，则需要用一个Recordset对象
		'并正确设置好CursorType和LockType为适当类型，然后调用Open方法打开
		Records.Open SqlStr, Cnn '使用这个语句,行数将返回-1,Set Records = Conn.Execute(SqlStr)
	End Function
	
'写入Excel表
	
	Function InputSheet(ByVal SheetName As String)
		Dim Columns, Rows As Integer
		Dim i, j As Integer
		
		Columns = Records.Fields.Count
		Rows = Records.RecordCount
		
		If Records.EOF = False And Records.BOF = False Then
		    For i = 0 To Rows - 1
		        For j = 0 To Columns - 1
		            Sheets(SheetName).Cells(i + 2, j + 1).Select
		            Sheets(SheetName).Cells(i + 2, j + 1) = Records.Fields.Item(j).Value
		        Next
		    Records.MoveNext
		    Next
		End If
		Sheets(SheetName).Cells(1, 1).Select
		MsgBox "Output!", vbOKOnly, "MySql to Excel"
	End Function
	
'把Excel写入MySql中的数据库

	Function InsertToMySql(ByVal SheetName As String, ByVal TblName As String)
		Dim SqlStr As String
		Dim i, j As Integer
		Dim Columns, Rows As Integer
		
		Columns = VBAProject.func_public.GetTotalColumns(SheetName)
		Rows = VBAProject.func_public.GetTotalRows(SheetName)
		
		Set Records = CreateObject("ADODB.recordset")
		'取得结果集并插入数据到数据库
		Set Records = CreateObject("ADODB.Recordset")
		'以下语句提供了插入思路，我只是把单条记录的插入方式改为循环，以把所有的记录添加到表中
		'rs.Open "insert   into   newtable  values('" & ActiveSheet.Cells(i, 1).Value & "'," & "'" & ActiveSheet.Cells(i, 2).Value & "')", cnn, 0
		For i = 2 To Rows
		    SqlStr = "INSERT INTO " & TblName & " values('" & Sheets(SheetName).Cells(i, 1).Value & "'" '注意：" values('"，字母“v”之前是有空格的！！！
		    For j = 2 To Columns
		        SqlStr = SqlStr & ",'" & Sheets(SheetName).Cells(i, j).Value & "'"
		    Next
		    SqlStr = SqlStr & ")"
		    Set Records = Cnn.Execute(SqlStr) 'rs.Open SqlStr, cnn, 0  不能用这条语句实现！！！
		Next
		MsgBox "Insert!", vbOKOnly, "Excel To MySql"
	End Function
	
'清除对象

	Function ClearObj()
		Set Cnn = Nothing
		Set Records = Nothing
	End Function
		
'获得数据表的字段名称
'OpenSchema可以获得数据库的各种信息

	Function InputColumns(ByVal SheetName As String)
		CnnOpen "localhost", "mydb", "employees", "root", ""
		Set Records = Cnn.OpenSchema(adSchemaColumns)
		
		Dim i As Integer
		i = 1
		While Not Records.EOF
		   Sheets(SheetName).Cells(1, i) = Records!COLUMN_NAME
		   i = i + 1
		   Records.MoveNext
		Wend
		
		CnnClose
		ClearObj
	End Function