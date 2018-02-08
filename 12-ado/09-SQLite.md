[http://club.excelhome.net/forum.php?mod=viewthread&tid=1381302&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=1381302&extra=page%3D1)


## 下面的代码会打开用”户数据”这个数据库文件，如果不存在会在这个路径上创建一个同名文件。 ##
	Sub 创建数据库()
		Dim conn As New ADODB.Connection '引用ADO
		Dim Connstr As String
		Connstr = "Driver={SQLite3 ODBC Driver};Database=" & ThisWorkbook.Path & "\用户数据.db"
		conn.Open Connstr
		conn.Close
	End Sub


## 下面的代码用于在数据库文件里创建一个表。 ##
	Sub 创建表()
		Dim conn As New ADODB.Connection '引用ADO
		Dim Connstr As String
		Connstr = "Driver={SQLite3 ODBC Driver};Database=" & ThisWorkbook.Path & "\用户数据.db"
		conn.Open Connstr
		conn.Execute "Create table 用户清单(用户编号,用户姓名,用户地址,联系电话)"  '在用户数据库文件下创建用户清单表，并创建4个字段名
		conn.Close
	End Sub

## 下面的代码是插入excel数据到清单数据表。 ##
	Sub 更新数据到DB()
		Dim conn As New ADODB.Connection '引用ADO
		Dim Connstr As String
		Connstr = "Driver={SQLite3 ODBC Driver};Database=" & ThisWorkbook.Path & "\用户数据.db"
		conn.Open Connstr
		conn.Execute "Create table 用户清单(ID,NAME,AGE,ADDRESS,SALARY)"  '在用户数据库文件下创建用户清单表，并创建5个字段名,如果表已存在会报错
		arr = [A1].CurrentRegion
		For i = 2 To UBound(arr)
		conn.Execute "Insert into 用户清单 values('" & Join(Application.Rept(Application.Index(arr, i, 0), 1), "','") & "')"
		Next
		conn.Close
	End Sub



## 查询数据基本就是select。和excel里面的SQL查询也是大体类似的。当然毕竟不同的软件，肯定语法子句有些差别。如下： ##
	Sub 查询数据()
		Dim conn As New ADODB.Connection '引用ADO
		Dim Connstr As String
		Connstr = "DSN=SQLite3 Datasource;Database=" & ThisWorkbook.Path & "\用户数据.db"
		conn.Open Connstr
		Set rst = conn.Execute("Select * from 用户清单 limit 2") 
		[A1].CopyFromRecordset rst
		conn.Close
	End Sub