[http://club.excelhome.net/thread-1159783-14-1.html](http://club.excelhome.net/thread-1159783-14-1.html)

# 处理table #
table数据处理，除了之前的两种通用方法外，还有以下几种方法：

## 1、html法 ##
   将table数据写入htmldocument对象，然后循环取出表格的各个元素。

-    优点：可以利用htmldocument对象整理表格。
-    缺点：需要学习html相关知识。

   以17楼作业二为例：

	Sub Main()
	    Dim strText As String
	    Dim arrData(1 To 1000, 1 To 3)
	    Dim i As Long, j As Long
	    Dim TR As Object, TD As Object
	    
	    With CreateObject("MSXML2.XMLHTTP")
	        .Open "POST", "http://www.pinble.com/Template/WebService1.asmx/Present3DList", False
	        .setRequestHeader "Content-Type", "application/json"
	        .Send "{pageindex:'1',lottory:'TC7XCData_jiangS',pl3:'',name:'江苏七星彩',isgp: '0'}"
	        strText = Split(JSEval(.responsetext), "<script")(0) '本例的script运行会提示错误，所以去除这部分script代码
	    End With
	    
	    With CreateObject("htmlfile")
	        .write strText
	        i = 0
	        For Each TR In .all.tags("table")(2).Rows
	            i = i + 1
	            j = 0
	            For Each TD In TR.Cells
	                j = j + 1
	                arrData(i, j) = TD.innerText
	            Next
	        Next
	    End With
	    
	    Set TR = Nothing
	    Set TD = Nothing
	    Cells.Clear
	    Range("C:C").NumberFormat = "@" '设置文本格式以显示数字前面的0
	    Range("a1").Resize(i, 3).Value = arrData
	End Sub
	
	Function JSEval(s As String) As String
	    With CreateObject("MSScriptControl.ScriptControl")
	        .Language = "javascript"
	        JSEval = .Eval(s)
	    End With
	End Function

## 2、QueryTable法： ##
   这个是excel自带的网抓利器。个人觉得它最大的优势就是处理table很方便。

   - 优点：处理table方便，代码简短。
-    缺点：会产生定义名称。多页循环时每页都会产生行字段名称，需要后续处理删除。

   仍以作业一的第1题为例：

	Sub Main()
	    Cells.Delete
	    With ActiveSheet.QueryTables.Add("url;http://data.bank.hexun.com/lccp/jrxp.aspx", Range("a1"))
	        .WebFormatting = xlWebFormattingNone '不包含格式
	        .WebSelectionType = xlSpecifiedTables '指定table模式
	        .WebTables = "2" '第2张table
	        .Refresh False
	    End With
	End Sub

代码相当简短。


## 3、复制粘贴法 ##：
   table部分的文字可以直接复制到单元格内，且保留数据原格式。

-    优点：只需取出table部分，不需分析数据内部结构。代码编写简便。
-    缺点：有时格式反而是累赘。

	Sub Main()
	    Dim strText As String
	    With CreateObject("MSXML2.XMLHTTP")
	        .Open "GET", "http://data.bank.hexun.com/lccp/jrxp.aspx", False
	        .Send
	        strText = .responsetext
	    End With
	    strText = "<table" & Split(Split(strText, "<table")(2), "</table>")(0) & "</table>"
	    CopyToClipbox strText
	    Cells.Clear
	    Range("a1").Select
	    ActiveSheet.Paste
	End Sub
	
	Sub CopyToClipbox(strText As String)
	    '文本拷贝到剪贴板
	    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
	        .SetText strText
	        .PutInClipboard
	    End With
	End Sub

小贴士：
点击Fiddler的Response框的WebView按钮可以看到HTML代码在网页上的显示效果。