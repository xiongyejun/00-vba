[[Environ$("comspec")](Environ$("comspec"))](http://club.excelhome.net/forum.php?mod=viewthread&tid=478544&extra=page%3D1)

# 在VBA中使用JAVASCRIPT和VBSCRIPT(1) #

javascript有许多函数和功能可以弥补VBA不足，如正则，数组，类，等等


## 1)以数组为例，用JAVASCRIPT排序 ##


	Sub fig8()
		Set x = CreateObject("msscriptcontrol.scriptcontrol")
		x.Language = "javascript"
		arr = Array("aa", "cc", "bb", "1a")
		kk = Join(arr, ",")
		x.addcode "function aa(bb){x=bb.split(',');x.sort();return x;}"
		cc = x.eval("aa('" & kk & "')")
		MsgBox cc
	End Sub

## 2)1)以数组为例，用JAVASCRIPT倒序 ##

	Sub fig8()	
		Set x = CreateObject("msscriptcontrol.scriptcontrol")
		x.Language = "javascript"
		arr = Array("aa", "cc", "bb", "1a")
		kk = Join(arr, ",")
		x.addcode "function aa(bb){x=bb.split(',');x.reverse();return x;}"
		cc = x.eval("aa('" & kk & "')")
		MsgBox cc	
	End Sub

## 用VBSCRIPT的简单例子 ##


	Sub fig8()
		Set x = CreateObject("msscriptcontrol.scriptcontrol")
		x.Language = "vbscript"
		
		x.addcode "sub aa(): msgbox ""hello.."":end sub "
		x.Run "aa"		
	End Sub

## 以前需要分开好几个模块，函数，现在可以统统放在一起了。。。。。 ##

	Sub fig8()
		
		Set x = CreateObject("msscriptcontrol.scriptcontrol")
		x.Language = "vbscript"
		
		x.addcode "sub aa(): msgbox ""hello.."":end sub : sub bb:msgbox 3:end sub :sub cc: msgbox ""cc"":end sub"
		x.Run "aa"
		x.Run "bb"
		x.Run "cc"
	
	End Sub

## 现在，代码也可以动态写入修改，不再需要考虑安全设置了哦 ##

	Sub fig8()
	
		Set x = CreateObject("msscriptcontrol.scriptcontrol")
		x.Language = "vbscript"
		
		arr6 = Array("aa", "bb", "cc")
		For Each arr In arr6
			x.addcode "sub " & arr & "(): msgbox """ & arr & "888"":end sub : "
			
			x.Run arr
		Next	
	
	End Sub

## 自定义函数的用法 ##

	Sub fig8()	
		Set x = CreateObject("msscriptcontrol.scriptcontrol")
		x.Language = "vbscript"		
		x.addcode "function sum(x,y):sum=x+y:end function "		
		bb = x.Run("sum", 2, 3)
		MsgBox bb	
	End Sub

## 动态改变窗口，文本框，单元格，range属性 ##

	'本例改[A1:z888]单元格为红色
	Sub fig88()
		Set X = CreateObject("msscriptcontrol.scriptcontrol")
		X.Language = "vbscript"
		X.addcode "Sub AA:XX.INTERIOR.COLORINDEX=3:End Sub "
		X.AddObject "XX", [A1:z888]
		X.Run "AA"	
	End Sub

## 设置和调用全局变量 ##

	Sub figvb()
		Set x = CreateObject("msscriptcontrol.scriptcontrol")
		x.Language = "vbscript"
		x.addcode "public x: sub aa(bb):x=bb*100:end sub"
		x.Run "aa", 3
		b = x.codeobject.x
		MsgBox b
	End Sub

## 代码放在单元格里不再是笑话：） ##

	Sub figvbs()
	    Set x = CreateObject("msscriptcontrol.scriptcontrol")
	    x.Language = "vbscript"
	    [a1] = "a1=3"
	    [a2] = "b1=4"
	    [a3] = "msgbox a1+b1"
	    For i = 1 To 3
		    x.executestatement Cells(i, 1)
	    Next
	End Sub


## 新建类可以不再需要类模块 ##

	Sub figvbs()
	    Dim x As Object
	    Dim rr As Object
	    
	    Set x = CreateObject("msscriptcontrol.scriptcontrol")
	    x.Language = "vbscript"
	    x.AddCode "Class AA:Public Sub Test():MsgBox ""类模块"":End Sub:End Class"
	    x.AddCode "Set YY=New AA"
	    Set rr = x.Eval("YY")
	    rr.Test
	End Sub

## 表达式可以直接拿来运算 ##

	Sub aa()
	    Dim x As Object
	    Dim Arr(2)
	    Dim kk As String
	    Dim bb
	    
	    Set x = CreateObject("msscriptcontrol.scriptcontrol")
	    x.Language = "vbscript"
	
	    Arr(0) = "3"
	    Arr(1) = "4*6"
	    Arr(2) = "SIN(5)"
	    kk = Join(Arr, "+")
	    x.ExecuteStatement ("MsgBox " & kk)
	    
	    kk = Join(Arr, "*")
	    bb = x.ExecuteStatement("MsgBox " & kk)
	End Sub

## msgbox ,inputbox 也可以作为变量 ##

	Sub figtest1()
	    Dim x As Object, i As Long
	    Dim aa As String, bb As String, kk As String
	    
	    Set x = CreateObject("msscriptcontrol.scriptcontrol")
	    x.Language = "vbscript"
	    aa = "msgbox "
	    bb = "cc=inputbox"
	    For i = 1 To 4
	        If i Mod 2 = 0 Then
	            kk = aa & "  " & i
	        Else
	            kk = bb & "(" & i & ")"
	        End If
	        x.executestatement (kk)
	    Next
	End Sub


## 数组也可以随意切割了 ##
	Sub JSArraySample()
	    Dim objJS As Object
	    Dim b As Object
	    Dim 文字列 As String
	    
	    Set objJS = CreateObject("ScriptControl")
	    With objJS
	        .Language = "JScript"
	        .AddCode "function JSSplit(s,d){return s.split(d);}"
	    End With
	  
	    文字列 = "a,b,c,d,e"
	  
	    Set b = objJS.CodeObject.JSSplit(文字列, ",")
		' '数组也可以随意切割了
	    MsgBox b.slice(0, 1)
	    MsgBox b.slice(1, 2)
	    MsgBox b.slice(2, 5)
	
	End Sub

## 功能更加强大的正则表达式 ##

	Sub figexp()
	    Dim js As Object
	    Dim Script As String
	    Dim result As String
	    
	    Set js = CreateObject("ScriptControl")
	    js.Language = "JScript"
	    Script = "'abcdefg'.match(/a/)"
	    result = js.eval(Script)
	    MsgBox result
	End Sub

## 学习研究了一下LDY版主的代码，发现一个现象， ##
jscript返回的对象应该是一个数组，可以在VB直接调用相关函数，但又可以直接显示所有元素

	Sub Mytest()
	    Dim sp1 As Object
	    Dim s As String
	    Dim aa
	    Dim bb As Object
	    
	    Set sp1 = CreateObject("ScriptControl")
	    sp1.Language = "JScript"
	    s = "function sortarr(arr){return arr.toArray();}"    '顺序
	    sp1.AddCode s
	    aa = Array("张", "王", "李", "赵", "钱", "孙", "周", "吴", "郑", "王")
	    Set bb = sp1.codeobject.sortarr(aa)
	    MsgBox bb
	    MsgBox bb.slice(1, 4)
	    MsgBox bb.concat("888").concat("777")
	    bb.push ("999")
	    MsgBox bb
	End Sub

## 数组可以直接合并 ##

	Sub Mytest()
	    Dim sp1 As Object
	    Dim s As String
	    Dim aa, aa2
	    Dim bb As Object, bb2 As Object, cc As String
	
	    Set sp1 = CreateObject("ScriptControl")
	    sp1.Language = "JScript"
	    s = "function sortarr(arr){return arr.toArray();}"    '顺序
	    sp1.AddCode s
	    aa = Array("张", "王", "李", "赵", "钱", "孙", "周", "吴", "郑", "王")
	    aa2 = Array("33", "王", "44")
	    Set bb = sp1.codeobject.sortarr(aa)
	    Set bb2 = sp1.codeobject.sortarr(aa2)
	    cc = bb.concat(bb2)
	    MsgBox cc
	End Sub

## 其他的强大的数组功能 ##

	Sub Mytest()
	    Dim sp1 As Object
	    Dim s As String
	    Dim aa
	    Dim bb As Object
	
	    Set sp1 = CreateObject("ScriptControl")
	    sp1.Language = "JScript"
	    s = "function sortarr(arr){return arr.toArray();}"    '顺序
	    sp1.AddCode s
	    aa = Array("张", "王", "李", "赵", "钱", "孙", "周", "吴", "郑", "王")
	    
	    Set bb = sp1.codeobject.sortarr(aa)
	    
	    bb.push ("999") '直接添加到数组末尾，不再需要重定义
	    MsgBox bb
	    
	    bb.unshift ("888") '直接添加到数组开头，不再需要重定义
	    MsgBox bb
	    
	    bb.pop '删除最后一个元素
	    MsgBox bb
	    
	    bb.shift '删除最前一个元素
	    MsgBox bb
	    
		'测试结果2,3并没有替换掉！
	    bb.splice 2, 3, "a", "b", "c" '直接替换数组

	    MsgBox bb
	End Sub

## 数组的读取 ##

	Sub Mytest()
	    Dim x As Object
	    Dim s As String
	    Dim i As Long
	    Dim y As Object
	
	    Set x = CreateObject("scriptcontrol")
	    x.Language = "jscript"
	    Set y = x.eval("aa=new Array()")
	    
	    For i = 1 To 100
	        y.push i
	    Next
	        
	    MsgBox x.eval("aa[" & 8 & "]")
	End Sub

## 把多维数组转换为一维 ##
	Sub Mytest()
	    Dim sc As Object
	    Dim a
	    Dim n As Object
	
	    [a1] = 1
	    [a2] = 2
	    [b1] = 3
	    [b2] = 4
	
	    Set sc = CreateObject("ScriptControl")
	    sc.Language = "JScript"
	    
	    a = [a1:b2]
	    
	    sc.AddCode "function aa(a){return new VBArray(a).toArray();}"
	    Set n = sc.CodeObject.aa(a)
	    
	    MsgBox n
	End Sub

## 获得当前屏幕的长宽，不用API ##
	Sub ava2()
	    Dim ie As Object
	    Dim win As Object
	    
	    Set ie = CreateObject("htmlfile")
	    Set win = ie.parentwindow
	    MsgBox win.screen.Width
	End Sub

## 把单元格作为对象传入js里 ##

	Sub ava2()
	    Dim x As Object
	    Dim y As Object
	    
	    Set x = CreateObject("scriptcontrol")
	    x.Language = "jscript"
	    
	    x.eval "function aa(aa) {return aa.value.toArray()}"
	    
	    Set y = x.Run("aa", [a1:b4])
	    MsgBox y
	
	End Sub

## 创建对象和属性 ##

	Sub ava2()
	    Dim x As Object
	    Dim y As Object
	    
	    Set x = CreateObject("scriptcontrol")
	    x.Language = "jscript"
	    x.eval "aa=new Object;aa.myname='fig7'"
	    Set y = x.eval("aa")
	    
	    MsgBox y.myname
	    
	    y.myname = "fig8"
	    MsgBox y.myname
	
	End Sub

## 调用其他模块，JAVASCRIPT 也可以有MSGBOX ##

注意把代码放入WORKBOOK,不能放入模块1，至于模块1应该传入什么对象，请版主和其他高手有空研究一下，谢谢

这里传入ME,就是thisworkbook，可以调用BB函数，在模块1我试过传入APPLICATION对象，但是不行，无法调用BB

如果用VBE.PROJECTS对象又要修改设置不太实用

	Private Sub kkk()
	    Dim m_sc As Object
	    
	    Set m_sc = CreateObject("ScriptControl")
	    With m_sc
	        .Language = "JScript"
	        .AddObject "o", Me
	        .EVAL "o.bb()"
	    End With
	   
	End Sub
	
	Public Sub bb()
	    MsgBox "kk"
	End Sub


## 多线程，同时运行，突破VBA程序运行单线程限制 ##
	Sub ava2()
	    Dim x As Object
	    Dim y, i
	    Dim ie As Object
	    
	    Set x = CreateObject("scriptcontrol")
	    Set ie = CreateObject("htmlfile")
	    x.Language = "jscript"
	    x.EVAL "var bb;function aa() {bb.range('a1')+=1;} ;function mm(cc,dd){bb=cc;dd.setInterval(aa,2000)}"
	    y = x.Run("mm", ActiveSheet, ie.parentWindow)
	    x.EVAL "var bb;function aa() {bb.range('a2')+=1;} ;function mm(cc,dd){bb=cc;dd.setInterval(aa,2000)}"
	    y = x.Run("mm", ActiveSheet, ie.parentWindow)
	    
	    For i = 1 To 888888888888#
	        [a3] = [a3] + 1
	        DoEvents
	    Next
	End Sub