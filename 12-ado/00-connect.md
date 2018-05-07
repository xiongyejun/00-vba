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



[http://club.excelhome.net/thread-1412070-1-1.html](http://club.excelhome.net/thread-1412070-1-1.html)

花了一些时间整理了在ADO中常用的连接字符串，方便查阅和比较。

并对其中的参数做了必要说明，期望让其能让多数人能看懂。

本帖内容涵盖了连接到Access，Excel，TXT，SQL Server，MySQL的连接字符串。

--------------------------------------------------------------------------------------------

## 1.Access  ##
Access 2003 Access 2007 Access 2010 Access 2013


### 本地文件： ###

Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\myFolder\myAccessFile.accdb;
Persist Security Info=False;

--------------------------------------------------------------------------------------------

### 网络文件（IP地址前为双反斜杠,例如：\\192.168.0.1\文件夹\文件.accdb）： ###

Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\server\share\folder\myAccessFile.accdb;

--------------------------------------------------------------------------------------------

带密码：

Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\myFolder\myAccessFile.accdb;
Jet OLEDB:Database Password=MyDbPassword;


--------------------------------------------------------------------------------------------

## 2.Excel ##
Excel 2003 Excel 2007 Excel 2010 Excel 2013

Excel 12.0 Xml中的后缀XML、MACRO可以省略

Xlsx文件

Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myExcel2007file.xlsx;
Extended Properties="Excel 12.0 Xml;HDR=YES";

Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myExcel2007file.xlsx;
Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1";

--------------------------------------------------------------------------------------------

Xlsb文件

Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myBinaryExcel2007file.xlsb;
Extended Properties="Excel 12.0;HDR=YES";

--------------------------------------------------------------------------------------------

Xlsm文件

Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myExcel2007file.xlsm;
Extended Properties="Excel 12.0 Macro;HDR=YES";

--------------------------------------------------------------------------------------------

Xls文件(Excel 97-2003)

Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myOldExcelFile.xls;
Extended Properties="Excel 8.0;HDR=YES";

Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyExcel.xls;
Extended Properties="Excel 8.0;HDR=Yes;IMEX=1";


参数说明

HDR=Yes：

这代表第一行是标题，不做为数据使用 ，如果用HDR=NO，则表示第一行不是标题，做为数据来使用。默认值YES

Excel 8.0：

对于Excel 97以上、2003及以下版本都用Excel 8.0，Excel 2007以上用Excel 12.0

IMEX(IMport EXport mode)：

IMEX是用来告诉驱动程序使用Excel文件的模式，其值有0、1、2三种，分别代表导出、导入、混合模式。当我

们设置IMEX＝1时将强制混合数据（数字、日期、字符串等）转换为文本。

但仅仅这种设置并不可靠，IMEX＝1只确保在某列前8行数据至少有一个是文本项的时候才起作用，它只是把查

找前8行数据中数据类型占优选择的行为作了略微的改变。例如某列前8行数据全为纯数字，那么它仍然以数字

类型作为该列的数据类型，随后行里的含有文本的数据仍然变空。另一个改进的措施是IMEX＝1与注册表值

TypeGuessRows配合使用，TypeGuessRows 值决定了ISAM 驱动程序从前几条数据采样确定数据类型，默认为“8

”。

可以通过修改“HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Excel”下的该注册表值来更改采

样行数，设置为0时表示采样所有行。

　　IMEX 三种模式：

　　当 IMEX=0 时为“汇出模式”(Export mode)，该模式开启的Excel档案只能用来做“写入”用途。
　　当 IMEX=1 时为“汇入模式”(Import mode)，该模式开启的Excel档案只能用来做“读取”用途。
　　当 IMEX=2 时为“连結模式”(Linked mode)，该模式开启的Excel档案支持“读取”和“写入”用途。


选择数据区域：

"SELECT [列名一], [列名二] FROM [表一$]"，Excel工作表名后面跟着一个“$”，并用[]括号括起来；如果

HDR=NO，也就是工作表没有标题，用F1，F2...引用相应的数据列。

"SELECT * FROM [Sheet1$a5:d10]"，选择A5到D10的数据区域。

数据区域也可以用Excel中定义的名称表示，假如有个工作簿作用范围的数据区名称datarange,查询语句为：

"SELECT * FROM [datarange]"

如果数据区名称作用范围是工作表，需要加上工作表名："SELECT * FROM [sheet1$datarange]"

有密保的工作簿：

如果Excel工作簿受密码保护，即使通过提供正确的密码与连接字符串，也无法打开它来进行数据访问。如果您

尝试打开，将收到以下错误信息：“无法解密文件”。

--------------------------------------------------------------------------------------------

## 3.文本文件 ##

分隔符列

Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\txtFilesFolder\;
Extended Properties="text;HDR=Yes;FMT=Delimited";

--------------------------------------------------------------------------------------------

定长列

Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\txtFilesFolder\;
Extended Properties="text;HDR=Yes;FMT=Fixed";

--------------------------------------------------------------------------------------------

FMT(Format) - 指定格式化类型。可以有如下值：

Delimited        文件被当做一个逗号分隔文件。逗号是默认分隔符。
Delimited(x)        文件被当做 'x’作为分隔符的文件
TabDelimited        文件被当做制表符分隔的文件
FixedLength        通过指定字段的固定长度来读取数据。

在注册表HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Text\可修改：
值为："Format" = "TabDelimited"或"Format" = "Delimited(;)"
如不指定，则为："Format"="Delimited( )"

--------------------------------------------------------------------------------------------

## 4.SQL Server ##

标准安全模式(Standard Security):

Provider=sqloledb;Data Source=myServerAddress;Initial Catalog=myDataBase;
User Id=myUsername;Password=myPassword;

--------------------------------------------------------------------------------------------

信任连接(Trusted connection):

Provider=sqloledb;Data Source=myServerAddress;Initial Catalog=myDataBase;
Integrated Security=SSPI;

参数说明：

Integrated Security（集成验证）为True时,连接语句中的UserID, Password是不起作用的，即采用windows身

份验证模式。只有设置为False或省略该项的时候，才按照UserID, Password来连接。Integrated Security可

以设置为: True, false, yes, no ，还可以设置为：SSPI ，相当于True.如果SQL SERVER服务器不支持这种方

式登录时，就会出错，你此时应使用SQL SERVER的用户名和密码进行登录，如： 

"Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=数据库名;
Data Source=192.168.0.1;User ID=sa;Password=密码"

integrated security=true表示以 Windows 身份验证的方式连接SQL。这种模式只允许SQL安装在本机上才能成

功登录。如果是远程登录模式，那么就应该使用用户名，密码的方式连接

Persist Security Info：是保存安全信息(密码)的，最好设置为false


--------------------------------------------------------------------------------------------

禁用连接池(Disable connection pooling):

Provider=sqloledb;Data Source=myServerAddress;Initial Catalog=myDataBase;
User ID=myUsername;Password=myPassword;OLE DB Services=-2;

--------------------------------------------------------------------------------------------

提示用户名和密码(Prompt for username and password)

oConn.Provider = "sqloledb"
oConn.Properties("Prompt") = adPromptAlways
oConn.Open "Data Source=myServerAddress;Initial Catalog=myDataBase;"


--------------------------------------------------------------------------------------------

通过IP地址连接(Connect via an IP address)


Provider=sqloledb;Data Source=190.190.200.100,1433;Network Library=DBMSSOCN;
Initial Catalog=myDataBase;User ID=myUsername;Password=myPassword;

"Network Library=DBMSSOCN"声明OLE DB使用TCP/IP替代Named  Pipes命名管道连接方式，
不加的话就使用MSSQL服务器端默认连接方式，不受程序控制。
支持的值包括：

                   dbnmpntw（命名管道）
                   dbmsrpcn（多协议，Windows RPC）
                   dbmsadsn (Apple Talk)
                   dbmsgnet (VIA)
                   dbmslpcn（共享内存）
                   dbmsspxn (IPX/SPX)
                   dbmssocn (TCP/IP)  
                   dbmsvinn (Banyan Vines)

--------------------------------------------------------------------------------------------

## 5.MySQL ##

Provider=MySQLProv;Data Source=mydb;User Id=myUsername;Password=myPassword;



MSSQL和MySQL使用的提供者驱动不同时，连接字符串都不会相同，不再列举。可在该网址查询：

https://www.connectionstrings.com/，基本上涵盖了所有类型的数据库连接字符串了！