[http://club.excelhome.net/forum.php?mod=viewthread&tid=1238755&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=1238755&extra=page%3D1)

有许多论友比较喜欢能自动关闭的Msgbox信息框，一般常见方式是调用Wscript.Shell对象的POPUP方法实现，如下面的代码：


    Dim wShell As Object
    Set wShell = CreateObject("Wscript.Shell")  '创建对象
    wShell.popup "执行完毕!", 2, "提示", 64      '执行popup方法，实现Msgbox信息框弹出
    Set wShell = Nothing                                '释放对象

但上面的方式可能有时会有些问题，比如：

一、有时会无法自动在指定的时间后自动自闭弹出的信息框；

二、在有些系统上可能会出现CreateObject("Wscript.Shell")失败而返回Nothing，这样的话信息框都不会弹出；

三、信息框弹出后，在信息框关闭前仍可以操作Excel中的工作簿窗体，在某些特定的情况可能会导致严重错误。

其实，要实现自动关闭的Msgbox只要调用API MessgeBoxTimeOut就可以很简单的实现了。顺便说下为什么VBA中没有一个类似MsgboxTimeOut的函数呢，这是因为VBA中的函数都是从VB6中继承来的，但是VB6生产于Windows 98时代，而MessgeBoxTimeOut这个API函数最早出现于Windows XP，所以VBA中自然就没有这个函数了。虽说现在VBA7版本出来了，但是它似乎仅仅是为了让VBA原有的功能可在后续的Windows版本中继续运行而已，故而没有新增什么东西。

	
	Option Explicit
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
	'>>>>>>>>   Author:     Joforn                            <<<<<<<<<<<<<<<<<<
	'>>>>>>>>   Email:      Joforn@sohu.com                   <<<<<<<<<<<<<<<<<<
	'>>>>>>>>   QQ:         42978116                          <<<<<<<<<<<<<<<<<<
	'>>>>>>>>   Last time : 10/31/2015                        <<<<<<<<<<<<<<<<<<
	'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
		
	#If VBA7 Then
	  Private Declare PtrSafe Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutW" ( _
	    ByVal hWnd As Long, ByVal lpText As LongPtr, _
	    ByVal lpCaption As LongPtr, ByVal wType As Long, _
	    ByVal wLange As Long, ByVal dwTimeout As Long) As Long
	#Else
	  Private Declare Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutW" ( _
	    ByVal hWnd As Long, ByVal lpText As Long, _
	    ByVal lpCaption As Long, ByVal wType As Long, _
	    ByVal wLange As Long, ByVal dwTimeout As Long) As Long
	#End If
	Private lngTimeOut As Long
	
	Public Property Let MsgboxTimeOutSecond(ByVal TimeOut As Long)
	  On Error GoTo LetSecondError
	  If TimeOut < 0 Then
	    lngTimeOut = 0
	  Else
	    lngTimeOut = TimeOut * 1000
	  End If
	  Exit Property
	LetSecondError:
	  lngTimeOut = &H7FFFFFFF
	End Property
	
	Public Property Let MsgboxTimeOut(ByVal TimeOut As Long)
	  If TimeOut < 0 Then
	    lngTimeOut = 0
	  Else
	    lngTimeOut = TimeOut
	  End If
	End Property
	
	Public Property Get MsgboxTimeOut() As Long
	  MsgboxTimeOut = lngTimeOut
	End Property
	
	Public Function Msgbox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
	                 Optional ByVal Title As String = vbNullString, Optional ByVal TimeOut As Long = -1&, _
	                 Optional ByVal LangeId As Long = 0&) As VbMsgBoxResult
	  'TimeOut以毫秒为单位，1 second = 1000 ms,TimeOut值为0时表示不自动返回,为负值时表示使用全局默认值
	  '如果信息框弹出后，用户未点击任何按钮，将返回3200,但如果Buttons的按钮值为VbOkOnly时，返回VbOk
	  
	  If TimeOut < 0 Then TimeOut = lngTimeOut
	  If Len(Title) < 1 Then Title = Application.Caption
	  Msgbox = MessageBoxTimeout(Application.hWnd, StrPtr(Prompt), StrPtr(Title), Buttons Or &H2000&, LangeId, TimeOut)
	End Function


说明：
一、MsgboxTimeOutSecond及MsgboxTimeOut两个属性只是为了方便大家设置全局默认自动关闭时间用的，这两个属性对应同一个值，但是是两个不同的单位：MsgboxTimeOut的单位值是毫秒，而MsgboxTimeOutSecond的单位是秒，这是为了方便有搞不清单位换算的筒子用的。但这两个属性设置的值只有在Msgbox省略TimeOut参数或是TimeOut参数值为负数时有效。为什么会添加这两个属性呢，主要是考虑到如果原有工程代码中有大量Msgbox要全部设置为自动关闭而增加的，因为有了它们，只要在工程运行的最开始处（比如：Workbook_Open事件处理过程）添加一条如MsgboxTimeOut = 1000这样的代码就可以轻松将所有的Msgbox指定为1秒后自动关闭，而不用再去修改原有代码；

二、Msgbox函数取消了原有系统自带Msgbox函数中的两个与帮助相关的参数(估计多数人都从来不用这个两参数，至少本人就极少用到^_^)；

三、本函数弹出的信息框样式是Windows 98样式，如果有不喜欢这个Style的筒子，请使用其它的方式来实现；

四、导入本模块后，可能会影响到其它的工作簿的Msgbox的Style，但不影响其正常功能；

五、如果你的程序将会在Windows 2000、98系统或是更低的Windows版本中运行，请不要使用本函数！！！