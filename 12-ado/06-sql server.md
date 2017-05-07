[http://club.excelhome.net/thread-583859-1-1.html](http://club.excelhome.net/thread-583859-1-1.html)

最近由于工作原因对excel vba+sql server 2000进行了学习，现将学习心得总结如下：

## 一、安装前准备工作： ##
在安装SQL server2000确保windowsXP安装了iis5.1。（http://tbg1.blog.163.com/blog/static/222360792009108105650566/）
## 二、SQL server2000安装： ##
SQL server2000安装（http://www.bitscn.com/os/windows/200604/5180.html）
## 三、远程连接SQL Server 2000服务器的解决方案： ##
（1）、ping服务器IP能否ping通
观察远程SQL Server 2000服务器的物理连接是否存在。如果不行，请检查网络，查看配置，当然得确保远程sql server 2000服务器的IP拼写正确。

（2） 在Dos或命令行下输入telnet 服务器IP 端口，看能否连通
如telnet 10.110.28.88 1433
通常端口值是1433，因为1433是SQL Server 2000的对于Tcp/IP的默认侦听端口。如果有问题，通常这一步会出问题。通常的提示是“……无法打开连接,连接失败"。

如果这一步有问题，应该检查以下选项。

1.检查远程服务器是否启动了sql server 2000服务。如果没有，则启动。

2.检查服务器端有没启用Tcp/IP协议，因为远程连接(通过因特网)需要靠这个协议。检查方法是，在服务器上打开 开始菜单->程序->Microsoft SQL Server->服务器网络实用工具，看启用的协议里是否有tcp/ip协议，如果没有，则启用它。

3.检查服务器的tcp/ip端口是否配置为1433端口。仍然在服务器网络实用工具里查看启用协议里面的tcp/ip的属性，确保默认端口为1433，并且隐藏服务器复选框没有勾上。

事实上，如果默认端口被修改，也是可以的，但是在客户端做telnet测试时，写服务器端口号时必须与服务器配置的端口号保持一致。如果隐藏服务器复选框被勾选，则意味着客户端无法通过枚举服务器来看到这台服务器，起到了保护的作用，但不影响连接，但是Tcp/ip协议的默认端口将被隐式修改为2433，在客户端连接时必须作相应的改变。

4.如果服务器端操作系统打过sp2补丁，则要对windows防火墙作一定的配置，要对它开放1433端口，通常在测试时可以直接关掉windows防火墙(其他的防火墙也关掉最好)。

5.检查服务器是否在1433端口侦听。如果服务器没有在tcp连接的1433端口侦听，则是连接不上的。检查方法是在服务器的dos或命令行下面输入netstat -a -n 或者是netstat -an，在结果列表里看是否有类似 tcp 127.0.0.1 1433 listening 的项。如果没有，则通常需要给sql server 2000打上至少sp3的补丁。其实在服务器端启动查询分析器，输入 select @@version 执行后可以看到版本号，版本号在8.0.2039以下的都需要打补丁。

## 四、excel与sql server服务器的连接： ##
在建立excel与sql server服务器的连接确保：工程--〉引用--〉选择“Microsoft ActiveX Data Objects 2.0(或者2.1-2.8)Library”

	Public Sub 服务器连接测试()
	    Dim cnn As New ADODB.Connection
	    Dim cnnStr As String
	    '建立与SQL Server数据库服务器的连接
	    Set cnn = New ADODB.Connection

	    cnn.ConnectionString = "Provider=SQLOLEDB; User ID=sa;Password =密码;Data Source=IP地址"
	    cnn.Open 
	   On Error GoTo 0
	    '判断数据库服务器连接是否成功
	    If cnn.State = adStateOpen Then
	        MsgBox "数据库服务器连接成功！", vbInformation, "连接服务器"
	    Else
	        MsgBox "数据库服务器连接失败！", vbInformation, "连接服务器"
	    End If
	    Set cnn = Nothing
	End Sub

## 五、在服务器上建立数据库用表： ##
	Public Sub 创建新的数据库及表()
	    Dim cnn As ADODB.Connection
	    Dim rs As ADODB.Recordset
	    Dim sql As String, mydata As String, mytable As String
	    mydata = "数据库名称"        '指定数据库名称
	   
	    '建立与SQL Server数据库服务器的连接
	    Set cnn = New ADODB.Connection
	    cnn.ConnectionString = "Provider=SQLOLEDB; User ID=sa;Password =123;password=密码;Data Source=IP地址"
	    cnn.Open
	    '判断数据库是否已经存在
	    sql = "select name from sysdatabases where name='" & mydata & "'"
	    Set rs = cnn.Execute(sql)
	    If rs.BOF = False Or rs.EOF = False Then
	        MsgBox "数据库<" & mydata & ">已经存在！请重新命名数据库！", vbCritical
	        Exit Sub
	    End If
	    '执行SQL语句创建数据库
	    sql = "create database " & mydata
	    cnn.Execute sql
	    MsgBox "数据库创建成功!", vbInformation, "创建数据库"
	    '关闭与SQL Server数据库服务器的连接
	    cnn.Close
	    '建立与刚刚创建的SQL Server数据库的连接
	    Set cnn = New ADODB.Connection
	    cnn.ConnectionString = "Provider=SQLOLEDB;" _
	        & "User ID=sa;" _
	        & "Password =密码;" _
	        & "Data Source=IP地址;" _
	        & "Initial Catalog=" & mydata
	    cnn.Open
	    '执行SQL语句创建数据表
	    sql = "create table 供货商信息" _
	        & "(供货商编码 varchar(10) not null,供货商名称 varchar(40) not null," _
	        & "通讯地址 varchar(30) not null,邮政编码 varchar(6) not null," _
	        & "联系电话 varchar(14) not null,传真号码 varchar(14) not null," _
	        & "联系人 varchar(10) not null,联系人电话 varchar(14) not null," _
	        & "联系人Email varchar(50) not null,备注 varchar(50))"
	    cnn.Execute sql
	    sql = "create table 物资信息" _
	        & "(物资类别 varchar(10) not null,物资编码 varchar(10) not null,物资名称 varchar(20) not null," _
	        & "规格型号 varchar(10) not null,单位 varchar(10) not null)"
	    cnn.Execute sql
	    MsgBox "数据表创建成功!", vbInformation, "创建数据表"
	    
	    cnn.Close
	    Set rs = Nothing
	    Set cnn = Nothing
	End Sub