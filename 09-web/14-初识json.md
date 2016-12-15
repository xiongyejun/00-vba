[http://club.excelhome.net/thread-1159783-20-1.html](http://club.excelhome.net/thread-1159783-20-1.html)

# 初识JSON #

JSON数据的特点：

- 1、用方括号扩住的是数组，数组内元素以逗号分隔。如：["甲","乙","丙"]、[1,2,3]
- 2、用花括号扩住的是对象，对象内各属性以逗号分隔，属性名和属性值以冒号分隔。同一对象里的属性名不会重复。如对象{"name":"甲","age":36}，含name、age两个属性，属性值分别为 “甲”和36。
- 3、对象的属性值可以是数组。数组的元素可以是对象。JSON数据就是数组对象嵌套的大集合。

比如，下面的JSON数据记录了甲乙二人的基本信息：


	[{"name":"甲","age":36,"children":[{"name":"甲儿","age":10},{"name":"甲女","age":7}]},{"name":"乙","age":28,"children":[{"name":"乙女","age":6}]}]

将其格式化后，看起来更清晰:


	[ 
	  {
	    "name" : "甲", 
	    "age" : 36, 
	    "children" : 
	    [
	      {
	        "name" : "甲儿", 
	        "age" : 10
	      },
	      {
	        "name" : "甲女", 
	        "age" : 7
	      }
	    ]
	  },
	  {
	    "name" : "乙", 
	    "age" : 28, 
	    "children" : 
	    [ 
	      {
	        "name" : "乙女", 
	        "age" : 6
	      }
	    ]
	  }
	]



解读：甲36岁，有10岁的甲儿和7岁的甲女两个子女；乙28岁，有6岁的乙女一个子女。

JSON数据易于解读，字符简短，方便网络传输，在网络中应用相当广泛。

敬请进入http://www.w3school.com.cn/json/index.asp学习更多JSON知识。

小贴士：
点击Fiddler的Response框的JSON按钮，可以很清晰的看到JSON数据结构。

# JSON转换成vba对象 #

1、JSON数组在vba内需要用For Each来获取其元素：（For Each 后面的变量不能定义为Object类型）

	Sub Test()
	    Const strJSON As String = "[""甲"",""乙"",""丙""]"
	    Dim objJSON As Object
	    Dim Cell '这里不能定义为object类型
	    With CreateObject("msscriptcontrol.scriptcontrol")
	        .Language = "JavaScript"
	        .AddCode "var mydata =" & strJSON
	        Set objJSON = .CodeObject
	    End With
	    Stop '查看vba本地窗口里objJSON对象以了解JSON数据在vba里的形态
	    For Each Cell In objJSON.mydata
	        Debug.Print Cell
	    Next
	End Sub

2、JSON对象在vba内可直接用“对象.属性”的方法获取，但当名称不被vba允许时，用CallByName函数获取：

	Sub Test()
	    Const strJSON As String = "{""name"":""甲"",""age"":36}"
	    Dim objJSON As Object
	    With CreateObject("msscriptcontrol.scriptcontrol")
	        .Language = "JavaScript"
	        .AddCode "var mydata=" & strJSON
	        Set objJSON = .CodeObject
	    End With
	    Stop '查看本地窗口
	    Debug.Print objJSON.mydata.age
	    Debug.Print objJSON.mydata.Name '此句出错
	End Sub

“name”作为vba对象的属性时会被自动首字母大写。而JavaScript里是区分大小写的，“Name”不能等同“name”，json数据里无“Name”属性，所以代码运行出错。

这时让我们请出“CallByName”吧。（使用前请先查看vba帮助中对此函数的说明）

将出错的那句代码改成：

	Debug.Print CallByName(objJSON.mydata, "name", VbGet)

ok。数据成功获取！

3、综合处理：

	Sub Test()
	    Const strJSON As String = "[{""name"":""甲"",""age"":36,""children"":[{""name"":""甲儿"",""age"":10},{""name"":""甲女"",""age"":7}]},{""name"":""乙"",""age"":28,""children"":[{""name"":""乙女"",""age"":6}]}]"
	    Dim objJSON As Object
	    Dim Person, Child
	    Dim arrData(1 To 100, 1 To 4)
	    Dim i As Long
	    
	    With CreateObject("msscriptcontrol.scriptcontrol")
	        .Language = "JavaScript"
	        .AddCode "var mydata =" & strJSON
	        Set objJSON = .CodeObject
	    End With
	    Stop '多多查看本地窗口，你才会进步更快。。。
	    
	    '为了编写方便，假设每个Person的children数组里至少有一个元素：
	    For Each Person In CallByName(objJSON, "mydata", VbGet)
	        For Each Child In CallByName(Person, "children", VbGet)
	            i = i + 1
	            arrData(i, 1) = CallByName(Person, "name", VbGet)
	            arrData(i, 2) = CallByName(Person, "age", VbGet)
	            arrData(i, 3) = CallByName(Child, "name", VbGet)
	            arrData(i, 4) = CallByName(Child, "age", VbGet)
	        Next
	    Next
	    Cells.Clear
	    Range("a1:d1").Value = Array("name", "age", "childname", "childage")
	    Range("a2").Resize(i, 4).Value = arrData
	End Sub

个人建议：为了区分vba对象本身的属性，建议JSON对象的属性都用CallByName表示。

更多在vba中使用JavaScript的例子，敬请参考figfig大师的帖子：

- http://club.excelhome.net/thread-478544-1-1.html
- http://club.excelhome.net/thread-483942-1-1.html
- http://club.excelhome.net/thread-484702-1-1.html


# 编写JavaScript代码处理JSON（一） #

个人为了练习分析JavaScript、JSON的能力，常编写JavaScript处理JSON成一个“表格文本”，放入剪贴板后粘贴到工作表内。

此法需要学习JavaScript知识。http://www.w3school.com.cn/js/index.asp

仍以之前的JSON数据为例：

	Sub Test()
	    Const strJSON As String = "[{""name"":""甲"",""age"":36,""children"":[{""name"":""甲儿"",""age"":10},{""name"":""甲女"",""age"":7}]},{""name"":""乙"",""age"":28,""children"":[{""name"":""乙女"",""age"":6}]}]"
	    Dim strJS As String
	    Dim strTable As String
	    
	    '为了编写方便，假设每个Person的children数组里至少有一个元素：
	    strJS = "var mydata=" & strJSON _
	        & ";var i,j;var s='name\tage\tchildname\tchildage\r';" _
	        & "for(i=0;i<mydata.length;i++){for(j=0;j<mydata[i].children.length;j++)" _
	        & "{s+=mydata[i].name+'\t'+mydata[i].age+'\t'+mydata[i].children[j].name+'\t'+mydata[i].children[j].age+'\r'}};"
	    strTable = JSEval(strJS)
	    Debug.Print strTable
	    
	    CopyToClipbox strTable
	    Cells.Clear
	    Range("a1").Select
	    ActiveSheet.Paste
	End Sub
	
	Function JSEval(strJS As String) As String
	    With CreateObject("MSScriptControl.ScriptControl")
	        .Language = "javascript"
	        JSEval = .Eval(strJS)
	    End With
	End Function
	
	Sub CopyToClipbox(strText As String)
	    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
	        .SetText strText
	        .PutInClipboard
	    End With
	End Sub

JavaScript的'\t'相当于 vbTab，'\r'相当于 vbCr，取出值后，每列之间用'\t'连接，每行之间用'\r'连接，这样的文本复制到剪贴板后，可直接粘贴到工作表形成表格数据。

本法非主流，不喜勿用！

# 编写JavaScript代码处理JSON（二） #

以70楼的获取QQ群成员列表的数据为例，再上一个处理JSON的例子：

	Sub Main()
	    Const gc As String = "" '群号
	    Const bkn As String = "" '从fiddler中获取
	    Const uin As String = "" 'QQ号
	    Const skey As String = "" '从fiddler中获取
	    Dim strText As String
	    Dim strJS As String
	    Dim strTable As String
	    
	    With CreateObject("WinHttp.WinHttpRequest.5.1")
	        .Open "GET", "http://qinfo.clt.qq.com/cgi-bin/qun_info/get_group_members_new?gc=" & gc & "&bkn=" & bkn, False
	        .setRequestHeader "Cookie", "uin=o" & uin & "; skey=" & skey
	        .Send
	        strText = .responsetext
	    End With
	    
	    strJS = "var mydata=" & strText _
	        & ";var x,qq,s='qq号\t昵称\t群名片\t等级\r';" _
	        & "var mems=mydata.mems,cards=mydata.cards,lv=mydata.lv,lvname=mydata.levelname;" _
	        & "for(x in mems){qq=mems[x].u;s+=qq+'\t'+mems[x].n+'\t'+(cards[qq]||'')+'\t'+lvname['lvln'+lv[qq].l]+'\r'}"
	    strTable = JSEval(strJS)
	
	    CopyToClipbox strTable
	    Cells.Clear
	    Range("a1").Select
	    ActiveSheet.Paste
	End Sub

此JSON数据内的很多对象都以QQ号为属性名，如mems[qq]可得该成员基本信息，cards[qq]得到群名片信息。所以可以遍历mems对象后获取qq号和昵称，再获取该qq群名片信息和等级信息组成table表。

如果不用JavaScript语句处理此JSON，因为群名片和群员等级在不同的对象里，取出后需要按照QQ号用字典来定位匹配。这样代码就相对较长了。有兴趣的朋友可以写个练练手。