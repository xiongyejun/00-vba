[http://club.excelhome.net/forum.php?mod=viewthread&tid=1107871&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=1107871&extra=page%3D1)

# 前言 #

Windows系统里提供了一个叫ScriptControl的OCX组件,我们可以用这个组件来实现脚本故事世界的精彩。

显示声明该组件：

vbe中引用 Microsoft Script Control 1.0

	Dim oMsp As MSScriptControl.ScriptControl    ''''声明对象变量
	Set oMsp = New MSScriptControl.ScriptControl    ''''实例化对象

该对象中本帖用到的属性、方法

- AddCode 方法 往脚本引擎中加入要执行的脚本
- Eval 方法 表达式求值
- CodeObject 属性 脚本暴露给宿主调用的对象。只读。（对象类型 ）
- Language 属性 设置或获取脚本引擎解释的语言，例如：VBScript、JavaScript。

## JavaScript ##
JavaScript 是属于网络的脚本语言！

JavaScript 是因特网上最流行的脚本语言。

明确一个概念，JavaScript不是java，也可以说两者基本没什么关系

JavaScript语言中有数组对象的概念，这个数组对象提供了很多方便我们处理数据的方法 

本帖正是因为这些方法的方便性而发的，目的是提供另一种解决问题的思路，仅供参考

本帖用到的方法

- splice(start,n,item1,item2..) 从start开始删去n位，并在该位置插入后面的元素。后面的元素可选，返回值为被删掉的n位
- sort() 进行升序排序。这个排序是基于Unicode的。sort(sortfunction) 使用数字排序的时候需要填入参数


另外：

JavaScript中toArray函数方法返回一个由 VBArray 转换而来的标准 JScript 数组。

本帖示例中用该方法的目的是将vba语音建立的数组转换成JavaScript的数组对象，这样才可以使用对象的方法 

必须注意的是经toArray转换后的数组都是一维的，就是说一个二维数组经toArray转换后返回的是个一维数组

其实vba中使用js还有很多意想不到之处，本人水平有限只能介绍这么多，不对的地方望大家斧正！！

	Sub 另类排序去重复()
	    Dim ojs As Object, m$
	    Set ojs = CreateObject("msscriptcontrol.scriptcontrol")
	    ojs.Language = "javascript"
	   ojs.AddCode "function y(z){x=z.split(',');x.sort(function(a,b){return a-b});for(i = 0; i<x.length;i++){if(x[i]==x[i+1]){x.splice(i, 1);i=i-1;}};return x;}"
	    
	    ''''生产测试用字符串
	    For i = 0 To 100
	        m = m & "," & Int(Rnd * 1000)
	    Next
	    MsgBox ojs.eval("y('" & Mid(m, 2) & "')")
	End Sub


	Sub 对单元格区域排序()
	    Set ojs = CreateObject("msscriptcontrol.scriptcontrol")
	    ojs.Language = "javascript"
	    ojs.AddCode "function sortarr(arr){a=arr.toArray();a.sort(function(a,b){return a-b});return a;}"
	    
	    aa = Range("a1:b2").Value
	    MsgBox ojs.codeobject.sortarr(aa)
	End Sub

	Sub 对单元格区域排序_去重复()
	    Set ojs = CreateObject("msscriptcontrol.scriptcontrol")
	    ojs.Language = "javascript"
	   ojs.AddCode "function sortarr(arr){x=arr.toArray();x.sort(function(a,b){return a-b});for(i = x.length; i>0;i--){if(x[i]==x[i-1]){x.splice(i, 1);}};return x;}"
	    
	    aa = Range("a1:b2").Value
	    MsgBox ojs.codeobject.sortarr(aa)
	End Sub
	
示例文件中的代码

	Sub 对单元格区域排序()
	    Dim oMsp As MSScriptControl.ScriptControl    ''''声明对象变量
	    Dim oJst As Object
	    Dim arr As Variant
	    
	    Set oMsp = New MSScriptControl.ScriptControl    ''''实例化对象
	    With oMsp
	        .Language = "javascript"
	        .Timeout = 60000
	        .AddCode "function sortarr(arr){x=arr.toArray();x.sort(function(a,b){return a-b});return x;}"
	    End With
	    
	    arr = Range("A1:B65536").Value
	    Set oJst = oMsp.CodeObject.sortarr(arr)
	    Range("A1").Resize(65536, 1).Value = Application.Transpose(Split(oJst.slice(0, 65536), ","))
	    Range("B1").Resize(65536, 1).Value = Application.Transpose(Split(oJst.slice(65536), ","))
	    
	End Sub
	
	Sub 对单元格区域排序_去重复()
	    Dim oMsp As MSScriptControl.ScriptControl    ''''声明对象变量
	    Dim oJst As Object
	    Dim arr As Variant, i&
	    
	    Set oMsp = New MSScriptControl.ScriptControl    ''''实例化对象
	    With oMsp
	        .Language = "javascript"
	        .Timeout = 60000
	        .AddCode "function sortarr(arr){x=arr.toArray();x.sort(function(a,b){return a-b});for(i = x.length; i>0;i--){if(x[i]==x[i-1]){x.splice(i, 1);}};return x;}"
	    End With
	    
	    arr = Range("A1:B65536").Value
	    Set oJst = oMsp.CodeObject.sortarr(arr)
	    Columns("A:B").ClearContents
	    arr = Split(oJst.slice(0, 65536), ",")
	    i = UBound(arr)
	    Range("A1").Resize(i + 1, 1).Value = Application.Transpose(arr)
	    arr = Split(oJst.slice(65536), ",")
	    i = UBound(arr)
	    If i > 0 Then Range("B1").Resize(i + 1, 1).Value = Application.Transpose(arr)
	End Sub