

    http://club.excelhome.net/thread-1100924-1-1.html

ado能连html的，但你这个文件不对，你这个文件原后缀名应该是html吧？你自己手工改成xls的。这个文件只是些框架代码，不是真正的数据文件。框架代码最后有一句：<frame src="1234.files/sheet001.htm" name="frSheet">
表示数据文件在1234.files文件夹下的sheet001.htm里。
以下代码测试成功：

	Sub Test()
	    Dim objCnn As Object
	    Dim objRst As Object
	    Set objCnn = CreateObject("Adodb.Connection")
	    Set objRst = CreateObject("Adodb.RecordSet")
	    objCnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
	                & "Extended Properties='HTML Import;HDR=YES;IMEX=1';" _
	                & "Data Source=" & ThisWorkbook.Path & "\1234.files\sheet001.htm"
	    Set objRst = objCnn.Execute("select * from [Table]")
	    objRst.Close
	    objCnn.Close
	    Set objRst = Nothing
	    Set objCnn = Nothing
	End Sub

有Title标签的html文档，要select from [title名]，比如你这个就是select * from [RTF Template]
没有title标签的，就select * from [table]

如果有多张table，就在后面加序号：
select * from [RTF Template] //这个是调取第一张table；
select * from [RTF Template1]//这个是调取第二张table；
。。。

可以用 Set objRst = objCnn.OpenSchema(20, Array(Empty, Empty, Empty, "TABLE")) 语句来查看表结构。
