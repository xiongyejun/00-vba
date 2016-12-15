[http://club.excelhome.net/thread-1159783-5-1.html](http://club.excelhome.net/thread-1159783-5-1.html)

# Cookie。。。 #

33楼的例子里，Cookie值的最初来源，来自服务器返回的回应头信息的“Set-Cookie”，它对winhttp对象的Cookie进行了设置。这种Cookie来源，不需要写setRequestHeader来模拟Cookie，只需发送一下相关页面的请求即能获取该Cookie。

但有些网页的Cookie值，来源并非是回应头信息的“Set-Cookie”。它可能是responsetext中的一部分。这时就只能setRequestHeader了。

（刚刚想找一个我以前写过的Cookie存在于responsetext的例子，发现网站竟然改版了，也“Set-Cookie”了！！）
征求Cookie存在于responsetext的例子。。。。

模拟Cookie的宗旨是向服务器发送一个它认可的Cookie。不管这个Cookie是回话Cookie还是持久Cookie，甚至该Cookie根本不存在你的电脑中（无论内存还是硬盘），只要服务器认可就行。

举例：
网站：http://club.excelhome.net/
操作：登录EH，在论坛首页点击“我发布的主题”，获取自己的主题帖数据。

EH的登录添加了验证码，这对我们初学网抓的人来说用xmlhttp/winhttp模拟登录EH是个无法克服的难题。

但是我们可以模拟Cookie来实现查询效果呀！

步骤：
1、抓包确认数据页面所在session。
2、点击该session的Request框的Cookie按钮可查看发送头内的Cookie内容。（挤在一起还是看不清楚）

![](http://files.c.excelhome.net/forum/201410/22/110759j4i4a5ko7vwvaxoz.png)

3、点击Request框的Header按钮查看Cookie更清晰：

![](http://files.c.excelhome.net/forum/201410/22/1108003u5khmxdnujk5ud5.png)

抓到的Cookie有很多，但不是每一个都需要模拟的。

经验：10位数字值一般都是用来记录时间的，这样的Cookie值一般不大会作为服务器的辨识标志，所以我们也一般不用模拟它。 （注意我用了“一般”这个词）

去除值为数字或空白的Cookie，只取值为字母数字混合物的Cookie写入代码中：

	Sub Main()
	    Dim strText As String
	    Const saltkey As String = "oUuXXXX"'请复制你自己的Cookie粘贴到这里。下同
	    Const sid As String = "tXXXX"
	    Const auth As String = "a30eEZTXXXXXXXXXXXXXXXXXXXX"
	    Const cookiereport As String = "f1fXXXXXXXXXXXXXXXXXXXXXXXX"
	    Const ulastactivity As String = "84cXXXXXXXXXXXXXXXXXXXX"
	    Const touclick As String = "70a9vPXXXXXXXXXXXXXXXXXXXX"
	    Const member_login_uid As String = "218917"
	    Const member_login_sid As String = "tXXXX"
	    
	    With CreateObject("WinHttp.WinHttpRequest.5.1")
	        .Open "GET", "http://club.excelhome.net/home.php?mod=space&do=thread&view=me ", False
	        .setRequestHeader "Cookie", _
	        "5WOj_b676_saltkey=" & saltkey _
	        & ";5WOj_b676_sid=" & sid _
	        & ";5WOj_b676_auth=" & auth _
	        & ";5WOj_b676_cookiereport=" & cookiereport _
	        & ";5WOj_b676_ulastactivity=" & ulastactivity _
	        & ";5WOj_b676_touclick=" & touclick _
	        & ";5WOj_b676_member_login_uid=" & member_login_uid _
	        & ";5WOj_b676_member_login_sid=" & member_login_sid
	        .Send
	        strText = .responsetext
	        Debug.Print strText
	    End With
	End Sub

运行。。查看立即窗口，获取数据成功。

继续测试，去除一些对获取数据无影响的Cookie（连论坛ID都去除了，大跌眼镜），代码最后成型为：
	
	Sub Main()
	    Dim strText As String
	    Const saltkey As String = "oUuXXXX"
	    Const auth As String = "a30eEZTXXXXXXXXXXXXXXXXXXXX"
	    
	    With CreateObject("WinHttp.WinHttpRequest.5.1")
	        .Open "GET", "http://club.excelhome.net/home.php?mod=space&do=thread&view=me ", False
	        .setRequestHeader "Cookie", _
	        "5WOj_b676_saltkey=" & saltkey _
	        & ";5WOj_b676_auth=" & auth
	        .Send
	        strText = .responsetext
	        Debug.Print strText
	    End With
	End Sub

这两个Cookie值就是EH的关键登录Cookie。它不需要保存在你的电脑里（本地电脑清除cookie、缓存后代码依旧运行成功）。我不知道它在服务器上的生存期有多长，我用了三天依然有效，哪怕在这三天内已经无数次的在不同地点手工重新登录过EH。有兴趣的话可以自己测试下它的生存期。
