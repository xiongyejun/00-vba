Microsoft.ACE.OLEDB.12.0 安装文件：



2007 Office system 驱动程序：数据连接组件
http://www.microsoft.com/downloads/details.aspx?displaylang=zh-cn&FamilyID=7554f536-8c28-4598-9b72-ef94e038c891

Microsoft Access 2010 数据库引擎可再发行程序包
https://www.microsoft.com/zh-cn/download/details.aspx?id=13255

https://www.connectionstrings.com/
[enter link description here](https://www.connectionstrings.com/)
# 使用OLEDB数据提供程序#

##访问XLSM:##

    Provier=Microsoft Ace OLEDB.12.0;Data Source=带路径的文件名;Extended Properties="Excel 12.0 Macro;HDR=NO";

##访问文本文件(如txt,csv等),数据源只要求指明文件所在的文件夹:##

    Provier=Microsoft Ace OLEDB.12.0;Data Source=文本文件所在文件夹;Extended Properties="text;HDR=YES;FMT=Delimited";


##访问数据Access##
 
  没有密码:
  
    Provier=Microsoft Ace OLEDB.12.0;Data Source=带路径的文件名;
    
  有密码:
   
   Provier=Microsoft Ace OLEDB.12.0;Data Source=带路径的文件名;Jet OLEDB:Database Password=密码;

##访问非Office数据库(如SQL Server/Oracle/MySQL)情况特别复杂,请参考提供商的说明.##

   举一例要求提供服务提供者的例子,使用SQL Server Native Client Provider提供者访问SQL Server 2012:
   
    Provider=SQLXMLOLEDB.4.0;Data Provider=SQLNCLI11;Data Source=服务器地址;Initial Catalog=数据库;User Id=账户名;Password=密码;

#使用OLEDB数据提供程序——无DSN连接#

##使用Ace OLEDB12.0的ODBC驱动程序访问Access2010数据库:##

	Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=数据库路径名;Uid=Admin;Pwd=;

##使用Microsoft SQL Server ODBC Driver访问SQL Server 2000##

	 Driver={SQL Server};Server=服务器地址;Database=数据库;Uid=用户名;Pwd=密码;



#使用OLEDB数据提供程序——有DSN连接#

##访问EXCEL连接字符串:##
     DSN=DSN名;
     
  ##访问带密码的Access数据库:##
  
     DSN=DSN名;UId=Admin;Pwd=;
    
   说明:
   
DSN名,是ODBC驱动程序的别名,有时也叫DSN数据源.是用户用ODBC管理器配置时自定义的.在配置中如果已经指定了数据源路径和文件名,则连接字符串中可不用带数据源路径名,也可带数据源路径名或其变量; 如果没有指定,则必须带数据源路径文件名.形如:

	    DSN=DSN名;Dbq=数据源路径名;


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