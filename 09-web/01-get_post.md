[http://club.excelhome.net/thread-1159783-3-1.html](http://club.excelhome.net/thread-1159783-3-1.html)

# 获取数据 #
## 直接获取-GET ##
xmlhttp/winhttp对象的Open方法的第一参数主要有两个值：GET 和 POST。（必须大写）

如何知道应该用GET还是POST呢？很简单，看之前用fiddler的数据网页的Request框Raw里的内容

## 直接获取-POST ##
Open第一参数是“POST"的时候，Send方法一般会有参数。

    Sub Main()
    	Dim strText As String
	    With CreateObject("MSXML2.XMLHTTP") 'CreateObject("WinHttp.WinHttpRequest.5.1")
	        .Open "POST", "http://cn.zso8.com/odds/search/", False
	        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			'.setRequestHeader "Referer", ""
	        .Send "type=2&CompanyID=11%7C%E6%BE%B3%E9%97%A8&leagueID=36&teamID=0&kind=1&port=&odds1=&do0=%E7%A1%AE%E5%AE%9A"
	        strText = .responsetext
	        Debug.Print strText
	    End With
    End Sub
经验：“POST”时，一般都需要模拟setRequestHeader的“Content-Type”字段（Header）

----------

小贴士：
Content-Type后面的值大部分都是application/x-www-form-urlencoded，但也有些网页不是这样的内容。每次代码模拟一定要和Request框内的内容保持一致，不要掉以轻心。

## 直接获取-静态参数 ##
在Request框内经常能看到各种参数。

参数用于客户端与服务器的交互。
参数的结构形式是：参数名1=参数值1&参数名2=参数值2&参数名3=参数值3.....

参数存在于SendData中（即vba代码里Send方法后面的那部分，又称POSTData、SendBody等），也存在于URL中。

在15楼的例子，参数存在于SendData中：
type=2&CompanyID=11%7C%E6%BE%B3%E9%97%A8&leagueID=36&teamID=0&kind=1&port=&odds1=&do0=%E7%A1%AE%E5%AE%9A

14楼的作业，参数存在于URL里。

点击Request框的WebForms按钮可清楚的查看各参数：
上面的“QueryString”框内显示的是URL部分的参数明细；
下面“Body"框内显示的是SendData部分的参数明细。

通过多次对网站不同操作的抓包对比，可分析出每个参数各自对应的页面的选项。据此，在代码里定义多个变量，可动态获取各种不同的查询结果。

没有值的参数，大多时候可以省略。比如：
type=2&CompanyID=11%7C%E6%BE%B3%E9%97%A8&leagueID=36&teamID=0&kind=1&port=&odds1=&do0=%E7%A1%AE%E5%AE%9A
可以省略写成：
type=2&CompanyID=11%7C%E6%BE%B3%E9%97%A8&leagueID=36&teamID=0&kind=1&do0=%E7%A1%AE%E5%AE%9A


上述URL或SendData里的参数都是大部分都是静态的（14楼作业2里有一个参数非静态），也就是说，不管何时何地，只要服务器后台代码不改变，相同的页面选择，参数总也是相同的。

但很多网页都有动态参数。动态参数有防盗链的效果。即使你操作手法完全相同，动态参数也会有不同。这个在之后的知识点里会有具体讲解。

小贴士：查询的数据有多页时，这个页码会以参数的形式体现在发包头内。

## 直接获取-转码 ##
在15楼的例子中：

从Request框的WebForms中可以看到：
CompanyID参数的值是“11|澳门”
do0参数的值是“确定”

但在Raw里看到的是：
CompanyID=11%7C%E6%BE%B3%E9%97%A8
do0=%E7%A1%AE%E5%AE%9A

这样的转换，可以用Javascript转换函数encodeURI或encodeURIComponent来实现（两者区别请自行百度）。

为了使代码容易解读，也为了让代码更换参数更加方便，我们将15楼的代码改动为：

    Sub Main()
	    Dim strText As String
	    With CreateObject("MSXML2.XMLHTTP") 'CreateObject("WinHttp.WinHttpRequest.5.1")
	        .Open "POST", "http://cn.zso8.com/odds/search/", False
	        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			'.setRequestHeader "Referer", ""
	        .Send "type=2&CompanyID=" & encodeURI("11|澳门") & "&leagueID=36&teamID=0&kind=1&port=&odds1=&do0=" & encodeURI("确定")
	        strText = .responsetext
	        Debug.Print strText
	    End With
    End Sub
自定义函数encodeURI请从11楼里复制。

但字符编码的类型有很多。上述只是一种比较常见的类型。一般来说，当一个中文转成3个16进制的码时（“%XX”为一个16进制码），适合用javascript的encodeURI或encodeURIComponent函数来解码。（实质是utf-8编码）

有时候是这样的编码：“abc一二三”转换为“abc%D2%BB%B6%FE%C8%FD” (ET的登录，用户名就是这样的编码）
在vba里，hex(asc("一"))="D2BB"，所以，用循环及文本函数可以很容易写出转码的自定义函数：

    Function GBKEnCode(strText)
	    Dim i, s
	    For i = 1 To Len(strText)
	        s = Hex(Asc(Mid(strText, i, 1)))
	        If Len(s) = 4 Then s = Left(s, 2) & "%" & Right(s, 2)
	        GBKEnCode = GBKEnCode & "%" & s
	    Next
    End Function

有人说，你这样连“abc”三个字符都转成“%61%62%63”了。。。没错，确实全部转了。但全部转了也可以Send成功的。不信你试试？

至于其他类型的编码，我想，只要你熟悉循环和文本函数，再了解下vba里的hex，asc，chr，ascw，chrw，总能写出自定义的转码函数的。

## 防盗链的处理-模拟Referer ##

当你的vba代码中：
GET部分后面的URL与Request框中的一摸一样；
或是POST的URL、SendData，与Request框中的一摸一样，Content-Type也没有忘记setRequestHeader；
但数据仍旧无法获取时，那你要获取的数据网页一定加了某种防盗链的措施。

此时要开始我们漫长的抓包调试过程了。（路漫漫其修远兮。。。）

调试这样的防盗链网页，如果在Request框中看到有这样的字样：
Referer: http://.......
那么，你首选模拟Referer（因为模拟Referer最简单）

在xmlhttp对象里模拟Referer是无效的。必须要用winhttp对象。


举例：
网站：http://www.sgs.gov.cn/lz/etpsInfo.do?method=index
操作：点选“企业名称”，在空白框内输入“通用技术集团医药控股有限公司”后点击“搜索”，在随后跳出的页面里点击“详细信息”，需要获取该公司的详细信息数据。

根据本楼之前的方法编写的代码：

    Sub Main()
	    Dim strText As String
	    With CreateObject("MSXML2.XMLHTTP") 'CreateObject("WinHttp.WinHttpRequest.5.1")
	        .Open "POST", "http://www.sgs.gov.cn/lz/etpsInfo.do?method=viewDetail", False
	        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			'.setRequestHeader "Referer", ""
	        .send "etpsId=150000012002040300047"
	        strText = .responseText
	        Debug.Print strText
	    End With
    End Sub


![](http://files.c.excelhome.net/forum/201410/21/213022rcnl0okcj47l987c.png)

模拟Referer后的代码：

    Sub Main()
	    Dim strText As String
	    With CreateObject("WinHttp.WinHttpRequest.5.1") 'CreateObject("MSXML2.XMLHTTP") '
	        .Open "POST", "http://www.sgs.gov.cn/lz/etpsInfo.do?method=viewDetail", False
	        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	        .setRequestHeader "Referer", "http://www.sgs.gov.cn/lz/etpsInfo.do?method=doSearch"
	        .send "etpsId=150000012002040300047"
	        strText = .responseText
	        Debug.Print strText
	    End With
    End Sub

再次运行，立即窗口即出现相关企业的详细信息。

![](http://files.c.excelhome.net/forum/201410/21/214127i8m06tvi6ji6kj7z.png)

小贴士：除了Content-Type、Referer和Cookie，其他setRequestHeader的情况很少见。

## 防盗链的处理-模拟Cookie ##


----------

**再次提醒，抓包前请先清除浏览器所有Cookie、缓存、历史数据。**

----------


举例：
网站：http://www.gzgczj.com:8080/costRegulatory/user.do?method=changeIndex&fareaId=1
操作：点击“控制价备案”，获取工程名称和招标控制价信息数据。

此网页模拟Referer后仍然无法获取数据。于是考虑模拟Cookie。
复制数据页面的cookie值，填入代码中（先不要去掉Referer的模拟）
![](http://files.c.excelhome.net/forum/201410/22/06385741c354wphxur44rw.png)

    Sub Main()
	    Dim strText As String
	    With CreateObject("WinHttp.WinHttpRequest.5.1") 
	        .Open "GET", "http://www.gzgczj.com:8080/costRegulatory/project.do?method=showProjectList&isVisitor=1&f_id=11011&t1413902083242", False
	        .setRequestHeader "Referer", "http://www.gzgczj.com:8080/costRegulatory/user.do?method=changeIndex&fareaId=1"
	        .setRequestHeader "Cookie", "E0685A9F6B708A1F1039BF2322B82A35"
	        .Send
	        strText = .responsetext
	        Debug.Print strText
	    End With
    End Sub

运行，，立即窗口显示“请重新登录”！

![](http://files.c.excelhome.net/forum/201410/22/063857saoabammmphb7abz.png)

这是为什么呢？这是因为Cookie失效了！我们需要模拟“新鲜”的Cookie才行！

模拟“新鲜”Cookie的方法：
1、复制上面Cookie值字符串。
2、搜索Session框内第一次出现此Cookie值的网页

![](http://files.c.excelhome.net/forum/201410/22/063857lhi55xllhi6rrsrw.png)

![](http://files.c.excelhome.net/forum/201410/22/063857emre7f0779ne55on.png)

3、编写代码，获取该Cookie，然后模拟写入数据网页：

    Sub Main()
	    Dim strText As String
	    Dim strCookie As String
	    
	    With CreateObject("WinHttp.WinHttpRequest.5.1")
	        .Option(6) = False ' 禁止重定向，以获取原网页信息
	        .Open "GET", "http://www.gzgczj.com:8080/costRegulatory/user.do?method=changeIndex&fareaId=1", False
	        .Send
	        strText = .getAllResponseHeaders '获取所有的回应头信息
	        Debug.Print strText: Stop '在立即窗口里查看头信息
	        strCookie = Split(Split(strText, "Set-Cookie: ")(1), ";")(0) '取出Cookie值
	    End With
	    
	    '在同一个winhttp对象里能保留cookie，为了体现设置cookie的作用，启用一个新的winhttp对象
	    With CreateObject("WinHttp.WinHttpRequest.5.1")
	        .Open "GET", "http://www.gzgczj.com:8080/costRegulatory/project.do?method=showProjectList&isVisitor=1&f_id=11011&t1413902083242", False
	        .setRequestHeader "Referer", "http://www.gzgczj.com:8080/costRegulatory/user.do?method=changeIndex&fareaId=1"
	        .setRequestHeader "Cookie", strCookie '模拟Cookie
	        .Send
	        strText = .responsetext
	        Debug.Print strText
	    End With
    End Sub

运行。。查看立即窗口，数据获取成功。
注意：本cookie是存在于重定向之前的网页的，要取到该cookie，必须禁用重定向。关于重定向，请参考108楼的帖子
不要用getResponseHeader("Set-Cookie")获取回应头信息中的Cookie值。如果有多个Cookie值，用getResponseHeader("Set-Cookie")获取不到完整的Cookie值。

你如果认为模拟Cookie只能写setRequestHeader的话，你就大错特错了。

xmlhttp对象和winhttp对象均可以保留Cookie。
也就是说，要获取某个网页的Cookie，可以先发送对此网页的请求（GET或POST）。只要xmlhttp对象或winhttp对象不关闭，此Cookie在对象运行期间全程有效。

据此，上述代码可以简化为：

    Sub Main()
	    Dim strText As String
	    With CreateObject("WinHttp.WinHttpRequest.5.1")
	        .Open "GET", "http://www.gzgczj.com:8080/costRegulatory/user.do?method=changeIndex&fareaId=1", False
	        .Send '此次send是为了获取cookie
	        .Open "GET", "http://www.gzgczj.com:8080/costRegulatory/project.do?method=showProjectList&isVisitor=1&f_id=11011&t1413902083242", False
	        .setRequestHeader "Referer", "http://www.gzgczj.com:8080/costRegulatory/user.do?method=changeIndex&fareaId=1"
	        .Send
	        strText = .responsetext
	        Debug.Print strText
	    End With
	End Sub

运行。。查看立即窗口，数据同样获取成功。

小贴士：用setRequestHeader方法模拟Cookie时，要用winhttp对象。发送网页的请求获取Cookie则xmlhttp、winhttp都可以。