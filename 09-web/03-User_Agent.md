[http://club.excelhome.net/thread-1159783-6-1.html](http://club.excelhome.net/thread-1159783-6-1.html)

# 获取数据-防盗链的处理-模拟User-Agent #

很少遇到需要模拟User-Agent的网页。

服务器可以根据发送头里的User-Agent辨别你是用手机还是电脑，是用IE浏览器还是用火狐浏览器。

举例：
EH查看他人的帖子（主题或回复）有些限制，需要登录才能查看所有会员的主题或回复。
但这种限制仅仅是在电脑上，手机不在此例。
因此，可以模拟User-Agent，伪装成在手机上浏览EH网站，查看他人帖子。

我们在电脑上也可以利用fiddler伪装成手机哦！
步骤：
1、打开Fiddler，勾选“Rules”-“User-Agents”-“WinPhone7”

![](http://files.c.excelhome.net/forum/201410/22/152711qm0y0iqiqxqayrqq.png)

2、打开浏览器，打开EH论坛，点击任一用户名，点击“回帖数”。

![](http://files.c.excelhome.net/forum/201410/22/152711ho5mmo3btkivx7t8.png)

![](http://files.c.excelhome.net/forum/201410/22/152711w40ttqltl3cbmlmw.png)

3、在Fiddler里查找数据网页，复制User-Agent后的字符串，写入代码。

![](http://files.c.excelhome.net/forum/201410/22/15271247nt2idagtdaagmg.png)

代码如下：

	Sub Main()
	    Dim strText As String
	    With CreateObject("MSXML2.XMLHTTP") 'CreateObject("WinHttp.WinHttpRequest.5.1")
	        .Open "GET", "http://club.excelhome.net/home.php?mod=space&uid=218917&do=thread&view=me&type=reply&from=space&mobile=yes", False
	        .setRequestHeader "User-Agent", "Mozilla/4.0 (compatible: MSIE 7.0; Windows Phone OS 7.0; Trident/3.1; IEMobile/7.0; SAMSUNG; SGH-i917)"
	        .Send
	        strText = .responsetext
	        Debug.Print strText
	    End With
	End Sub