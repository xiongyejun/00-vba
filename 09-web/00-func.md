[http://club.excelhome.net/forum.php?mod=viewthread&tid=1159783&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=1159783&extra=page%3D1)

# vba网抓常用方法： #

## 1、xmlhttp/winhttp法： ##
用xmlhttp/winhttp模拟向服务器发送请求，接收服务器返回的数据。
优点：效率高，基本无兼容性问题。
缺点：需要借助如fiddler的工具来模拟http请求。

## 2、IE/webbrowser法： ##
创建IE控件或webbrowser控件，结合htmlfile对象的方法和属性，模拟浏览器操作，获取浏览器页面的数据。
优点：这个方法可以模拟大部分的浏览器操作。所见即所得，浏览器能看到的数据就能用代码获取。
缺点：各种弹窗相当烦人，兼容性也确实是个很伤脑筋的问题。上传文件在IE里根本无法实现。（有实现方法？请一定告诉我）

## 3、QueryTables法： ##
因为它是excel自带，所以勉强也算是一种方法。其实此法和xmlhttp类似，也是GET或POST方式发送请求，然后得到服务器的response返回到单元格内。
优点：excel自带，可以通过录制宏得到代码，处理table很方便。代码简短，适合快速获取一些存在于源代码的table里的数据。
缺点：无法模拟referer等发包头（如果你有在QT中模拟referer的方法，请一定告诉我）

本帖主要讲述的是第一种方法。


常用代码及自定义函数：

## 1、网抓主体代码： ##

    Sub Main()

	    Dim strText As String
	
	    With CreateObject("MSXML2.XMLHTTP") 'CreateObject("WinHttp.WinHttpRequest.5.1")'
	        .Open "POST", "", False
	        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"	
	        .setRequestHeader "Referer", ""	
	        .Send	
	        strText = .responsetext	
	        Debug.Print strText	
	    End With

    End Sub

代码里的很多""就是留给你的填空题。。。
xmlhttp/winhttp对象的属性和方法可以网上百度学习（不学也暂时影响不大），内容不多。

## 2、Javascript表达式求值： ##

    Function JSEval(strText  As String) As String

	    With CreateObject("MSScriptControl.ScriptControl")
	        .Language = "javascript"
	        JSEval = .Eval(strText)
	    End With

    End Function

## 3、url转码： ##


    Function encodeURI(strText As String) As String

	    With CreateObject("msscriptcontrol.scriptcontrol")
	        .Language = "JavaScript"
	        encodeURI = .Eval("encodeURIComponent('" & strText & "');")
	    End With

    End Function

javascript提供了六个转码函数：
escape，unescape，encodeURI，encodeURIComponent，decodeURI，decodeURIComponent
具体用法请百度。我只能说我最常用的是encodeURIComponent。

## 4、流数据转成指定编码的文本： ##

    Function ByteToStr(arrByte, strCharset As String) As String
	    With CreateObject("Adodb.Stream")	
	        .Type = 1 'adTypeBinary	
	        .Open	
	        .Write arrByte	
	        .Position = 0	
	        .Type = 2 'adTypeText	
	        .Charset = strCharset	
	        ByteToStr = .Readtext	
	        .Close	
	    End With

    End Function

## 5、文本按指定编码转为流数据： ##

    Function StrToByte(strText As String, strCharset As String)

	    With CreateObject("adodb.stream")	
	        .Mode = 3 'adModeReadWrite	
	        .Type = 2 'adTypeText	
	        .Charset = strCharset	
	        .Open	 
	        .Writetext strText	
	        .Position = 0	
	        .Type = 1 'adTypeBinary	
	        '.Position = 2 '保留BOM头则不需此行代码，去除三个字节的BOM头就填入3，去除两个字节的就填入2	
	        StrToByte = .Read	
	        .Close	
	    End With

    End Function

注：某些文本转为流后，前面会添加几个字节的BOM头，用来被某些软件识别是什么编码。如UTF-8编码的前面有三个字节的BOM头，Unicode前面有两个字节的BOM头。大家可以视情况选择保留或去除这些BOM头。

## 6、二进制流转成文件： ##

    Sub ByteToFile(arrByte, strFileName As String)

	    With CreateObject("Adodb.Stream")	
	        .Type = 1 'adTypeBinary	
	        .Open	
	        .Write arrByte	
	        .SaveToFile strFileName, 2 'adSaveCreateOverWrite	
	        .Close	
	    End With

    End Sub

## 7、文本拷贝到剪贴板： ##


    Sub CopyToClipbox(strText As String)'文本拷贝到剪贴板

	    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")	
	        .SetText strText	
	        .PutInClipboard	
	    End With

    End Sub

呃，原谅我吧，我是个怀旧的人。我的机器的配置目前为止仍旧是32位的winXP+2003office，IE也刚升级到IE8，之前一直用的IE6。弦月大师（xmyjk）所点评的内容我没有办法提供呀。。甩泪。。。

----------
突然想起HtmlWindow也可以直接执行Javascript函数得出值。
用64位office的朋友可以测试一下下面的代码能不能通过：
替代上面自定义函数2的：

    Function EvalByHtml(strText As String) As String

	    With CreateObject("htmlfile")	
	        .write "<html><script></script></html>"	
	        EvalByHtml = CallByName(.parentwindow, "eval", VbMethod, strText)	
	    End With

    End Function

替代上面自定义函数3的：

    Function encodeURIByHtml(strText As String) As String

	    With CreateObject("htmlfile")	
	        .write "<html><script></script></html>"	
	        encodeURIByHtml = CallByName(.parentwindow, "encodeURIComponent", VbMethod, strText)	
	    End With

    End Function

给Dom添加一个空的script就可以直接执行js函数了，非常好用。

为防止vba自动篡改大小写，把js函数名作为文本放在callbyname的参数里。


[http://club.excelhome.net/thread-1159783-36-1.html](http://club.excelhome.net/thread-1159783-36-1.html)

	Sub 转码示例01()
	    Debug.Print escape("搜房网")'输出结果：%u641C%u623F%u7F51
	    Debug.Print ChtoJ3("搜房网")'输出结果：\u641c\u623f\u7f51
	
	    Debug.Print unescape("%u641C%u623F%u7F51")'输出结果：'搜房网
	    Debug.Print unescape("\u641c\u623f\u7f51")'输出结果：搜房网
	End Sub
	
	Function escape(strInput As String) As String
	    With CreateObject("msscriptcontrol.scriptcontrol")
	        .Language = "JavaScript"
	        escape = .Eval("escape('" & strInput & "');")
	    End With
	End Function
	
	Function unescape(strTobecoded As String) As String
	    With CreateObject("msscriptcontrol.scriptcontrol")
	        .Language = "JavaScript"
	        unescape = .Eval("unescape('" & strTobecoded & "');")
	    End With
	End Function
	
	Function ChtoJ3(szCode As String)
	    With CreateObject("MSScriptControl.ScriptControl")
	        .Language = "JavaScript"
	        .addcode "function decode(str){return escape(str).replace(/%/g,'\\')}"
	        ChtoJ3 = .Eval("decode('" & szCode & "')")
	    End With
	End Function


[http://www.soso.io/article/16029.html](http://www.soso.io/article/16029.html)

关于VBS采集，网上流行比较多的方法都是正则，其实 htmlfile 可以解析 html 代码，但如果 designMode 没开启的话，有时候会包安全提示信息。

但是开启 designMode (@预言家晚报 分享的方法) 的话，所有js都不会被执行，只是干干净净的dom文档，所以在逼不得已的情况下开启 designMode 一般情况保持默认即可。

	Sub test_htmlfile()
	    Dim HTML As Object
	    Dim http As Object
	    Dim strHtml As String
	    Dim post_list As Object
	    Dim el As Object
	    
	    Set HTML = CreateObject("htmlfile")
	    Set http = CreateObject("Msxml2.ServerXMLHTTP")
	     
	    HTML.designMode = "on" ' 开启编辑模式
	     
	    http.Open "GET", "http://www.cnblogs.com/", False
	    http.send
	    strHtml = http.responseText
	     
	    SetClipText strHtml
	    
	    HTML.write strHtml ' 写入数据
	    Set post_list = HTML.getElementById("post_list")
	    For Each el In post_list.Children
	      Debug.Print el.getElementsByTagName("a")(0).innerText
	    Next
	End Sub


`http://blog.csdn.net/tylm22733367/article/details/52596990`    

WinHttp.SetTimeouts 60000, 60000, 60000, 3000 '设置操作超时时间 

'WinHttp.SetTimeouts resolveTimeout, connectTimeout, sendTimeout, receiveTimeout
 
- 'resolveTimeout = 10000 '解析 DNS 名字的超时时间，10000 毫秒。 
- 'connectTimeout = 10000 '建立 Winsock 连接的超时时间，10000 毫秒。 
- 'sendTimeout = 120000 '发送数据的超时时间，120000 毫秒。 
- 'receiveTimeout = 60000 '接收 response 的超时时间，60000 毫秒。 


- WinHttp.Option(4) = 13056 '忽略错误标志 
- WinHttp.Option(6) = False '为 True 时，当请求页面重定向跳转时自动跳转，False 不自动跳转，截取服务端返回的302状态。 
- 
WinHttp.Open "GET", "http://www.baidu.com/", False 'GET 或 POST, Url, False 同步方式；True 异步方式 

## '组成 HTTP 头信息  ##
- WinHttp.SetRequestHeader "Accept", "*/*" '接受数据类型 
- WinHttp.SetRequestHeader "Accept-Language", "zh-cn,zh" '用户系统语言 
- WinHttp.SetRequestHeader "User-Agent", "Mozilla/6.0" '用户浏览器信息 
- WinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded" '编码方式 
- WinHttp.SetRequestHeader "Referer", "http://www.baidu.com/" '来路 
- WinHttp.SetRequestHeader "Connection", "Close" 'Close = 不保持连接，Keep-Alive = 保持连接(持久连接) 
- 'WinHttp.SetRequestHeader "Accept-Encoding", "gzip, deflate" '如果发送，会返回 gzip, deflate 压缩过的编码 
- 'WinHttp.SetRequestHeader "Content-Length", Len(Str) '内容长度，Post 方式用的。 
- WinHttp.SetRequestHeader "Cookie", "test cookie!" '设置 Cookie 
- 
## '发送数据  ##

WinHttp.Send 'Post 方式：WinHttp.Send (参数) 

WinHttp.WaitForResponse '等待返回请求，XMLHTTP中也可以使用 


## '输出结果  ##

- Response.Write WinHttp.Status '当前 HTTP 状态 
- 'Response.Write WinHttp.ResponseText '文本数据 
- Response.BinaryWrite WinHttp.ResponseBody '二进制数据流数据 


## Option()相关：  ##

- WinHttpRequestOption_UserAgentString = 0; 
- WinHttpRequestOption_URL = 1; 
- WinHttpRequestOption_URLCodePage = 2; 
- WinHttpRequestOption_EscapePercentInURL = 3; 
- WinHttpRequestOption_SslErrorIgnoreFlags = 4; 
- WinHttpRequestOption_SelectCertificate = 5; '13056 = 忽略错误标志 
- WinHttpRequestOption_EnableRedirects = 6; '为True时,当请求的页面中有跳转时,抓取跳转页面信息.False相反不抓取 
- WinHttpRequestOption_UrlEscapeDisable = 7; 
- WinHttpRequestOption_UrlEscapeDisableQuery = 8; 
- WinHttpRequestOption_SecureProtocols = 9; 
- WinHttpRequestOption_EnableTracing = 10; 
- WinHttpRequestOption_RevertImpersonationOverSsl = 11; 
- WinHttpRequestOption_EnableHttpsToHttpRedirects = 12; 
- WinHttpRequestOption_EnablePassportAuthentication = 13; 
- WinHttpRequestOption_MaxAutomaticRedirects = 14; 
- WinHttpRequestOption_MaxResponseHeaderSize = 15; 
- WinHttpRequestOption_MaxResponseDrainSize = 16; 
- WinHttpRequestOption_EnableHttp1_1 = 17; 
-  
- WinHttpRequestOption_EnableCertificateRevocationCheck = 18; 

## 方法 Description 说明  ##

- Abort 中止一个WinHTTP的 发送方法。 
- GetAllResponseHeaders 检索所有的HTTP响应头。 
- GetResponseHeader 检索HTTP响应头。 
- Open 打开一个HTTP连接到HTTP资源。 
- Send 发送一个HTTP请求到HTTP服务器。 
- SetAutoLogonPolicy 设置当前自动登录策略。 
- SetClientCertificate 选择一个客户端证书发送到一个安全的超文本传输协议（HTTPS）服务器。 
- SetCredentials 设置要使用的凭据与HTTP服务器，要么是原籍国或代理服务器。 
- SetProxy 设置代理服务器信息。 
- SetRequestHeader 添加，更改或删除一个HTTP请求标头。 
- SetTimeouts 指定以毫秒为单位，个人的时间超过了一个组件发送/接收操作。 
- WaitForResponse 指定的等待时间，以秒为一个异步传送方法来完成的值，SetTimeouts。 

## 该 WinHttpRequest 对象定义以下属性。 
## 属性 访问类型 说明 ## 

- Option 读/写 设置或检索一个WinHTTP的选项值。 
- ResponseBody 只读 检索作为无符号字节数组的响应实体机构。 
- ResponseStream 只读 检索机构作为响应实体的IStream。 
- ResponseText 只读 作为文本检索响应实体机构。 
- Status 只读 从上次检索响应的HTTP状态代码。 
- StatusText 只读 获取HTTP状态的文本。 

## Events 活动  ##

The WinHttpRequest object defines the following events. 
WinHttpRequest 对象定义的下列事件。 

Event 事件 Description说明 

- OnError Occurs when there is a run-time error in the application. 
- 当发生一个应用程序运行时错误时发生。 
- OnResponseDataAvailable Occurs when data is available from the response. 
- 当响应数据可用时发生。 
- OnResponseFinished Occurs when the response data is complete. 
- 响应数据完成时发生。 
- OnResponseStart Occurs when the response data starts to be received. 
- 开始收到响应数据时发生。 

- Remarks 备注 

- The WinHttpRequest object uses the IErrorInfo interface to provide error data. 
- 该 WinHttpRequest 对象使用 IErrorInfo 接口来提供错误数据。 
- A description and numerical error value can be obtained with the Err object in Microsoft Visual Basic Scripting Edition (VBScript), 
- and the Error object in Microsoft JScript. 
- The lower 16 bits of an error number correspond to the values found in Error Messages. 
- Note For Windows XP and Windows 2000, see Run-Time Requirements. 
- 注意：对于 Windows XP 和 Windows 2000，请参阅运行时间要求。 


## Requirements 要求  ##

- 最低支持的客户端 Windows XP, Windows 2000 Professional with SP3 Windows XP 
- 最低支持的服务器 Windows Server 2003, Windows 2000 Server with SP3 
- 
- 可再发行 WinHTTP 5.0 and Internet Explorer 5.01 or later on Windows XP and Windows 2000. 
- WinHTTP的5.0和Internet Explorer 5.01或更高版本的Windows XP和Windows 2000。 
- IDL HttpRequest.idl HttpRequest.idl 
- Library Winhttp.lib Winhttp.lib 
- DLL Winhttp.dll Winhttp.dll
