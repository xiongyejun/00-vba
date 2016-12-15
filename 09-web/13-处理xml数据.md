[http://club.excelhome.net/thread-1159783-17-1.html](http://club.excelhome.net/thread-1159783-17-1.html)

# 处理xml数据 #
8楼的例子返回的就是一个xml文档。

在excel里，有一个很快捷的方式导入xml文档：

	Sub Main()
	    ThisWorkbook.XmlImport _
	        URL:="http://www.cffex.com.cn/fzjy/tjsj/pztj/pzrtj/2014/index.xml", _
	        ImportMap:=Nothing, _
	        Overwrite:=True, _
	        Destination:=ActiveSheet.Range("a1")
	End Sub

如果需要自定义提取的内容，你或者可以这样写

	Sub Main()
	    Dim arrEM(1 To 4), arrEMname
	    Dim arrData(1000, 1 To 4)
	    Dim i As Long, j As Long
	    With CreateObject("MSXML2.XMLHTTP")
	        .Open "GET", "http://www.cffex.com.cn/fzjy/tjsj/pztj/pzrtj/2014/index.xml", False
	        .send
	        arrEMname = Array(, "productid", "tradingday", "volume", "openinterest")
	        With .responseXML 
	            For i = 1 To 4
	                Set arrEM(i) = .getElementsByTagName(arrEMname(i))
	            Next
	            For i = 0 To arrEM(1).Length - 1
	                For j = 1 To 4
	                    arrData(i, j) = arrEM(j)(i).Text
	                Next
	            Next
	        End With
	    End With
	    Cells.Clear
	    Range("a1:d1").Value = Array("品种", "日期", "总成交量", "总持仓量")
	    Range("a2").Resize(i, 4).Value = arrData
	End Sub

小贴士：
1、当URL是一个格式化的xml文档时，xmlhttp可将服务器返回的数据整理为一个DomDocument对象放在responseXML中。

你也可以CreateObject("MSXML2.DomDocument")然后用load(url)或者loadxml(text)方法获得DomDocument对象。

2、点击Fiddler的Response框的XML按钮，可以很清晰的看到XML数据结构。