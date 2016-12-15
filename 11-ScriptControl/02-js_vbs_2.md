[http://club.excelhome.net/forum.php?mod=viewthread&tid=483942&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=483942&extra=page%3D1)

# 在VBA中使用JAVASCRIPT和VBSCRIPT（2） #

## 介绍 JSON  ##

JSON(JavaScript Object Notation) 是一种轻量级的数据交换格式。易于人阅读和编写。同时也易于机器解析和生成。它基于JavaScript Programming Language（http://www.crockford.com/javascript）, Standard ECMA-262 3rd Edition - December 1999（http://www.ecma-international.or ... cma-st/ECMA-262.pdf）的一个子集。JSON采用完全独立于语言的文本格式，但是也使用了类似于C语言家族的习惯（包括C, C++, C#, Java, JavaScript, Perl, Python等）。这些特性使JSON成为理想的数据交换语言。

JSON建构于两种结构：

“名称/值”对的集合（A collection of name/value pairs）。不同的语言中，它被理解为对象（object），纪录（record），结构（struct），字典（dictionary），哈希表（hash table），有键列表（keyed list），或者关联数组 （associative array）。 

值的有序列表（An ordered list of values）。在大部分语言中，它被理解为数组（array）。 

这些都是常见的数据结构。事实上大部分现代计算机语言都以某种形式支持它们。这使得一种数据格式在同样基于这些结构的编程语言之间交换成为可能。

JSON具有以下这些形式：

对象是一个无序的“‘名称/值’对”集合。一个对象以“{”（左括号）开始，“}”（右括号）结束。每个“名称”后跟一个“:”（冒号）；“‘名称/值’ 对”之间使用“,”（逗号）分隔。 

数组是值（value）的有序集合。一个数组以“[”（左中括号）开始，“]”（右中括号）结束。值之间使用“,”（逗号）分隔。 

值（value）可以是双引号括起来的字符串（string）、数值(number)、 ture、false、 null、对象（object）或者数组（array）。这些结构可以嵌套。 

字符串（string）是由双引号包围的任意数量Unicode字符的集合，使用反斜线转义。一个字符（character）即一个单独的字符串（character string）。 

除去一些编码细节，以下描述了完整的语言。

字符串（string）与C或者Java的字符串非常相似。除去未曾使用的八进制与十六进制格式，数值（number）也与C或者Java的数值非常相似。

空白可以加入到任何符号之间。
	
	Sub figjson()
	    Dim x As Object, y As Object
	    Dim aa As String, s As String
	
	    aa = "{ ""myname"":""figfig"", ""myid"":""888"" }"
	    
	    Set x = CreateObject("ScriptControl")
	    x.Language = "JScript"
	       
	    s = "function j(s) { return eval('(' + s + ')'); }"
	    x.AddCode s
	    Set y = x.CodeObject.j(aa)
	      
	    MsgBox y.myname
	    MsgBox y.myid
	End Sub

## 例子2 ##
	Sub figjson()
        Dim x As Object, y As Object
        Dim aa As String, s As String
    
        aa = "{myname:""alonely"", age:24, email:[""aa4@bb.com"",""aa@gmail.com""], family:{parents:[""父亲"",""母亲""],toString:function(){return ""家庭成员"";}}}"
        Set x = CreateObject("ScriptControl")
        x.Language = "JScript"
           
        s = "function j(s) { return eval('(' + s + ')'); }"
        x.AddCode s
        Set y = x.Run("j", aa)
          
        MsgBox y.myname
        MsgBox y.age
        MsgBox VBA.CallByName(y, "email", VbGet)
        MsgBox y.family
        MsgBox y.family.parents
    End Sub


## 多重结构，树状显示，类似XML节点树，代码比XML简洁得多 ##
	Sub figjson()
	    Dim x As Object, y As Object
	    Dim aa As String, s As String
	
	    aa = "{""myname"":""Michael"",""myaddress"":{""city"":""Beijing"",""street"":"" Chaoyang Road "",""postcode"":100025}}"
	   
	    Set x = CreateObject("ScriptControl")
	    x.Language = "JScript"
	       
	    s = "function j(s) { return eval('(' + s + ')'); }"
	    x.AddCode s
	    Set y = x.Run("j", aa)
	      
	    MsgBox y.myname
	    MsgBox y.myaddress
	    MsgBox y.myaddress.city
	    MsgBox y.myaddress.postcode
	End Sub

## 数组放入对象里 ##
	Sub figjson()
	    Dim x As Object, y As Object
	    Dim aa As String, s As String
	    
	    aa = "{ ""people"": [{ ""firstName"": ""Brett"", ""lastName"":""McLaughlin"", ""email"": ""brett@newInstance.com"" },{ ""firstName"": ""Jason"", ""lastName"":""Hunter"", ""email"": ""jason@servlets.com"" }, { ""firstName"": ""Elliotte"", ""lastName"":""Harold"", ""email"": ""elharo@macfaq.com"" }]}"
	    Set x = CreateObject("ScriptControl")
	    x.Language = "JScript"
	       
	    s = "function j(s) { return eval('(' + s + ').people[1]'); }"
	    x.AddCode s
	    Set y = x.Run("j", aa)
	    MsgBox y.firstName
	    
	    MsgBox VBA.CallByName(y, "email", VbGet)
	End Sub

## 可用单引号代替2个双引号，简化写法，如例子一 ##
    Sub figjson()
        Dim x As Object, y As Object
        Dim aa As String, s As String
    
        aa = "{ 'myname':'figfig', 'myid':'888' }"
    
        Set x = CreateObject("ScriptControl")
        x.Language = "JScript"
       
        s = "function j(s) { return eval('(' + s + ')'); }"
        x.AddCode s
        Set y = x.CodeObject.j(aa)
      
        MsgBox y.myname
        MsgBox y.myid
    End Sub

## 在EXCEL中的应用 ##

如SHEET1 A1单元格为AA, B1单元格为BB

    Sub figjson()
        Dim x As Object, y As Object
        Dim aa As String, s As String
    
        Range("A1").Value = "aa"
        Range("B1").Value = "bb"
        aa = "{ '" & [a1] & "':'" & [b1] & " ' }"
    
        Set x = CreateObject("ScriptControl")
        x.Language = "JScript"
       
        s = "function j(s) { return eval('(' + s + ')'); }"
        x.AddCode s
        Set y = x.CodeObject.j(aa)
        MsgBox VBA.CallByName(y, "aa", VbGet)
    End Sub

## 传递数值值 ##

    Sub figjson()
        Dim x As Object, y As Object
        Dim aa As String, s As String
    
        Set x = CreateObject("ScriptControl")
        x.Language = "JScript"
       
        s = "var a=2 ;var b=3;var cc={a:a,b:b}"
        x.AddCode s
        Set y = x.CodeObject.cc
        MsgBox y.a
    End Sub

## 动态添加数据 ##
    Sub figjson()
        Dim x As Object, y As Object
        Dim aa As String, s As String
    
        Set x = CreateObject("ScriptControl")
        x.Language = "JScript"
       
        s = "var a=2 ;var b=3;var cc={a:a,b:b};cc['电话']=8888;"
        x.AddCode s
        Set y = x.CodeObject.cc
        MsgBox y.电话
    End Sub

## 数据动态变化 ##

    Sub figjson()
        Dim x As Object, y As Object
        Dim aa As String, s As String
    
        Set x = CreateObject("ScriptControl")
        x.Language = "JScript"
       
        s = "var a=2 ;var b=3;var cc={a:a,b:b};cc['电话']=8888;"
        x.AddCode s
        Set y = x.CodeObject.cc
        MsgBox y.电话
    
        s = "cc['电话']=9999;"
        x.AddCode s

        MsgBox y.电话
    End Sub

## 用变量来查询 ##
    Sub figjson()
        Dim x As Object, y
        Dim aa As String, s As String
        Dim kk As String
    
        Set x = CreateObject("ScriptControl")
        x.Language = "JScript"
       
        s = "var cc={name:'figfig',id:'888',tel:'1234'};"
        x.AddCode s
      
        kk = "name"
        y = x.eval("cc['" & kk & "']")
        MsgBox y
    
        kk = "id"
   
        y = x.eval("cc['" & kk & "']")
        MsgBox y
        
        kk = "tel"
        y = x.eval("cc['" & kk & "']")
        MsgBox y
    End Sub

## 与VB比较，代码更加简洁明了，可作为小型数据库 ##
    Sub figjson()
        Dim x As Object, Name As String
        Dim aa As String, Address As String
        Dim s As String
    
        Set x = CreateObject("ScriptControl")
        x.Language = "JScript"
       
        s = "var address={aa:'us',bb:'cn',cc:'uk'}"
        x.AddCode s

        Name = "bb"
        Address = x.eval("address['" & Name & "']")
        MsgBox Address
        
        Name = "cc"
        Address = x.eval("address['" & Name & "']")
        MsgBox Address
    End Sub

## 类似数组，可增加和删除数据 ##
    Sub figjson()
        Dim x As Object, Name As String
        Dim aa As String, Address As String
        Dim s As String, i As Long
    
        Set x = CreateObject("ScriptControl")
        x.Language = "JScript"
       
        s = "var address={bb:'0' };"
        x.AddCode s
          
        For i = 1 To 100
            s = "address[" & i & "]=" & i & ";"
            x.AddCode s
        Next
        
        Address = x.eval("address[88]")
        MsgBox Address
        
        Address = x.eval("address[77]")
        MsgBox Address
        
        x.eval ("delete address[77]")
        Address = x.eval("address[77]")
        MsgBox Address
    End Sub

