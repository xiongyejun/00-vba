[http://club.excelhome.net/thread-438643-1-1.html](http://club.excelhome.net/thread-438643-1-1.html)

# flash和office vba 交互问题的解决方案 #


ldhyob大侠老早以前发表过“借助FLASH技术美化VBA操作界面”(excelhome中：http://club.excelhome.net/viewth ... hlight=flash&page=1，officefans中：http://www.officefans.net/cdb/viewthread.php?tid=15539&highlight=%BD%E8%D6%FAFLASH%BC%BC%CA%F5%C3%C0%BB%AFVBA%B2%D9%D7%F7%BD%E7%C3%E6)的文章，

让OFFCE vba编程界面得到了很好的改观。

但是FLASH PLAYER8之后的版本和VBA交互往往不成功。问题根源就是Flash的FSCommand（）函数向vba发消息VBA接收不到。网上无数的网友都为这个问题困惑。一直以来没有这方面的需求，就没有去好好琢磨。最近由于工作中需要用到这方面的东西，所以抽空学习并试验了一下，FSCommand函数传递信息问题，引起的原因是：flash9和flash10的安全设置问题。
解决这个问题，我总结了一下有如下几个种解决方案：

## 方法1：通过flashplayer官方网站在线修改“全局安全设置”， ##

链接为：http://www.macromedia.com/suppor ... ings_manager04.html
在其中设置：“全局安全性设置”面板中选择“始终允许”（默认方式是始终询问）

![](http://files.c.excelhome.net/forum/month_0905/20090524_9ac4365b4117a32f3206uswe6XlqcR8L.jpg)

这种方式设置的好处是：可以保证不出错

这种方式设置的坏处是：比较麻烦，必须要上网才能修改。另外设置成“始终允许”存在一定安全隐患（当然这里也能修改信任位置，但总觉得比较麻烦）。

## 方法2：使用手动修改或者bat批处理方式设置 ##

建立批处理bat文件如下：
	
	echo off
	cls
	echo 设置相关运行环境
	pause
	c:
	cd %windir%\system32\Macromed\Flash
	md FlashPlayerTrust
	cd FlashPlayerTrust
	echo C:\ >myTrustFiles.cfg
	cd %userprofile%\Application Data\Macromedia\Flash Player\#Security
	md FlashPlayerTrust
	cd FlashPlayerTrust
	echo C:\ >myTrustFiles.cfg
	echo 设置完成。

将上述代码放到bat文件中，双击自动运行后即可

这样做，就必须要单独一个bat文件，往往使用起来不方便。于是有了方法3

## 方法3：通过VBA程序自动处理 ##

在thisworkbook中放入如下代码

	Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
	Private Sub myflashTrustFiles_creater()
	    Dim systemdisk As String
	    systemdisk = GetSysDir()
	    Dim yyy, a As Object
	    Set yyy = CreateObject("Scripting.FileSystemObject")
	    If yyy.FolderExists(systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust") = True Then yyy.DeleteFolder systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust"
	    yyy.CreateFolder (systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust")
	    If yyy.FileExists(systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust\myTrustFiles.cfg") = True Then
	        Kill systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust\myTrustFiles.cfg"
	    End If
	    Set yyy = CreateObject("Scripting.FileSystemObject")
	    Set a = yyy.CreateTextFile(systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust\myTrustFiles.cfg", True)
	    a.WriteLine (systemdisk & "\")
	    a.Close
	    If yyy.FolderExists(systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust\FlashPlayerTrust") = True Then yyy.DeleteFolder systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust\FlashPlayerTrust"
	    yyy.CreateFolder (systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust\FlashPlayerTrust")
	    
	    If yyy.FileExists(systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust\FlashPlayerTrust\myTrustFiles.cfg") = True Then
	        Kill systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust\FlashPlayerTrust\myTrustFiles.cfg"
	    End If
	    Set yyy = CreateObject("Scripting.FileSystemObject")
	    Set a = yyy.CreateTextFile(systemdisk & "\WINDOWS\system32\Macromed\Flash\FlashPlayerTrust\FlashPlayerTrust\myTrustFiles.cfg", True)
	    a.WriteLine (systemdisk & "\")
	    a.Close
	    Set yyy = Nothing
	    Set a = Nothing
	End Sub
	Private Function GetSysDir() As String
	    Dim sSave     As String
	    Dim Ret     As Long
	    sSave = Space(255)
	    Ret = GetSystemDirectory(sSave, 255)
	    GetSysDir = Left$(sSave, 2)
	End Function

然后在Workbook_Open中调用即可，即是插入如下代码：

    Call myflashTrustFiles_creater

## 方法4：使用webbrowser控件 ##

这种方法是

通过webbrowser控件方式调用swf文件，

其中flash点击按键弹出窗口：

	on (release) { 
	geturl("javascript:openwindow(’’http://www.webjx.com’’,’’’’,’’toolbars=no,location=no,scrollbars=no,status=no,resizable=no,width=500,height=500’’)") 
	}

然后在vba中通过BeforeNavigate2事件获得该url地址并分析处理，得到传递消息的目的；
附件（这是直接用了网上一位朋友的东西）

这种做法需要独立的swf文件（当然也可以在office以对象方式嵌入，然后通过vba提取出来，不过相当麻烦）


几种方法各有千秋，不过我更习惯方法三。


[http://club.excelhome.net/forum.php?mod=viewthread&tid=935035&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=935035&extra=page%3D1)

## 方法5：手动修改全局安全，完成flash与vba交互 ##

这是在没有系统权限的公司电脑上测试成功的方法：

### 1，右键单击flash播放界面 ###

![](http://files.c.excelhome.net/forum/201210/24/100131mscj2mr3hk82sred.png)

### 2，选择【高级】——【受信任位置设置】 ###

![](http://files.c.excelhome.net/forum/201210/24/1001536oe0ok3yifbobpj6.png)

### 3，【添加】-【添加文件夹】-【一定要选中C盘】【确定】 ###

![](http://files.c.excelhome.net/forum/201210/24/100233gjtugumgutmm3hz7.png)

vba与flash8以上的交互就这么解决了~~~~~（当初为了解决交互问题，呕心沥血啊！泪奔！）

解决了vba与flash交互问题 ，我们现在来看一个简单示例————【flash向vba传送信息】

只是简单的示例，这一步到达了，就入门了，后面就看你自己的开发了
