# 正则与VBA的交互—正则表达式的实现 #

在继续学习正则元字符特性或编制自己的正则表达式时,常常需要对其测试.你可以用一些专门的正则测试工具(推荐RegxBuddy);也可以自己编制VBA代码进行测试。不过建议初学者，经常编写VBA代码进行测试，这样可以提高今后实际应用正则的能力。所以，在进一步学习正则元字符特性之前，我们先介绍正则与VBA的交互的相关知识。你可以快速阅读或越过本章内容，在以后具体应用时，再经常回头查阅。当然也可以用上一章学到的知识详细研究本章内容,在以后的学习中专注于正则表达式本身.

用正则处理文本，是通过正则表达式与程序设计语言的交互来实现的。其交互方式在不同编程语言中分为三大类：

一是集成式。Perl语言本身内建正则操作符，可以直接作用于正则表达式.操作符作用于正则表达式就像数学的+-号作用于数字一样.不需要构建正则对象。例如:任务是要把变量$text保存的文本中的空行替换为标签（<P>）。

正则表达式

	^$         表示空行.

在Perl语言中,可以用一句代码实现替换：$text=~ s/^$/<p>/g

二是函数式处理。Java等语言,为正则处理提供了一些便捷函数,以节省工作量.用户不需要首先创建一个正则对象,而是用静态函数提的临时对象来完成正则处理,处理完后把临时对象抛弃.正则表达式对于函数相当于一个参数,这种方式的优点是”随手”可用,但不适宜在对时间要求很高的循环中使用.所以java也提供了下面讲到的面向对象的程序式处理.

三是面向对象的程序式处理。这是大多数编程语言的正则处理方式。VBA平台采用的也是这种方式。面向对象的程序式处理方式，首先必须创建一个正则对象的实例,然后设置对象必要的属性，最后用对象的方法来完成指定的任务。(提示:不同编程语言的正则对象具有的属性和方法，其项目多少或功能强弱有所不同，所以，在VBA中使用正则如果发现没有某种其它语言的方法或属性，请不要感到困惑)

在上一章中，我们给出了一个用VBA删除行尾空格的正则处理例子，它代表了一般的代码框架模式，下面再看一看它的结构特点,并对每一部分的代码段进行剖析：

	Sub test()
	    Dim regx,S$,Strnew$                                       1.定义变量代码段
	    S=”正则表达式其实很简单     “                   2.目标文本字串变量赋值代码段
	    Set regx=createobject(“vbscript.regexp”)   3.创建正则对象代码段
	    Regx.pattern=”\s+$”                                       4.设置正则对象的pattern属性代码段
	    Regx.global=true                                              5.设置正则对象的其它属性代码段
	    Strnew=regx.replace(s,””)                              6.应用正则对象方法代码段
	    Msgbox strnew                                                 7.处理返回值代码段
	End sub

1.定义变量代码段

不必讲解了吧.

2.目标文本字符串赋值代码段

目标文本,可能存在于文本文档、Word文档、HTML文档或Excel文档等文档之中。正则对象并不能直接作用于这些文档，只能作用于它们的副本。所以用VBA正则处理这些文档，必须首先从这些文档中读出字符串并赋值于字符变量。如果任务是修改文本,那么,你可能需要编写额外的代码将修改后的文本字符串重新写回原文档中.

例:假如目标文本存在于当前表格A1单元格中.可使用下列代码赋值于字符变量S
S=Activesheet.[a1]

目标文本也可能分别存在于一个数组中,那么,你可能需要通过循环逐一处理.

你也可以直接以输入的方式,赋值给字符变量,就像上面的例子.这时特别注意的是:半角双引号是VBA语言中的保留字符,如果目标文本中本身含有半角双引号,则必须转义,转义方法是:用重复的双引号表示一个双引号.

例:目标文本为:”我们用”汗牛充栋”、”学富五车”形容一个人读的书、拥有的知识多。”.

将之赋值给S的代码为：

S=”我们用””汗牛充栋””、””学富五车””形容一个人读的书、拥有的知识多。”

3.创建正则对象代码段

文本处理的各种操作,都是通过操作正则对象来完成的.所以必须创建正则对象.VBA创建或声明正则对象有两方式：早期绑定和后期绑定，你可以根据自己喜好选择其一：

早期绑定: （需要在VBE--工具--引用中勾选Microsoft VBScript Regular Expressions 5.5）

     Dim regx AS RegExp
     Set regx=new regexp (或dim regx as new regexp)

后期绑定: 

     Set regex = CreateObject("VBScript.RegExp")

利用上述两种方式创建或声明正则对象,实际上是调用Microsoft VBScript脚本的regexp正则对象。Microssoft VBScript脚本,包含在Internet Eeplorer 5.5以及之后的版本中.该脚本中的正则表达式执行的是ECMA-262第3版所规定的标准，与JavaScript脚本中的正则执行标准是相同的。1.0版只是为了向后兼容的目的,功能很弱。

(提示:在VBA中也可调用JavaScript(Jscript)或ruby等脚本中的正则对象,Jscript的元字符及特性与VBscript是一样的,但它的方法或属性要多一点,或者说对正则的支持更强一些.ruby本人不懂,不太了解它的元字符集,只是看到论坛上有人使用)

4.设置对象的pattern属性

语法:object.pattern=”正则表达式”

Object是一个正则对象.

把自己编制的正则表达式,以字符串的形式赋值给pattern属性。注意要用英文双引号将正则表达式包围起来.

并且要在对象名与属性名之间用英文点号隔开.属性名pattern是保留字,固定不变的,对象名是用户自定义的。

接下来的两个步骤是对正则对象的操作,通过设置或使用正则对象的属性和方法,以实现对文本的处理.正则对象的属性和方法不多,列表于下:

	
	属性			方法
	Global全局属性		test
	IgnoreCase大小写属性	replace
	MultiLine多行属性	Execute

5.设置对象的其它属性

除Pattern属性外,正则对象还有其它三个属性，其属性值有False和True，默认值都是False。如果要使用默认属性，可以不用显示设置；如果要改变默认属性，则需要显示设置:

Global：当属性值为False时,只要在目标文本中,找到一个匹配时,即停止搜索。如果想要找出目标文本中的所有匹配，那么需要把它的属性值设置为True。

IgnoreCase：设置对英文字母大小写是否敏感。默认值False, 对大小写敏感；设置为True,忽略大小写.

MultiLine：它影响且只影响元字符^和$的意义。值为False，无论目标文本是多少行，整个文本中则只有一个开始位置，^表示第一行的开始；只有一个行结束位置，$表示文本末尾位置。值为True，那么，^和$分别表示每一行的行首和行尾位置。

下面来完成一个简单的任务，再具体认识各属性的使用方法：

有一两行的文本：

- Aaa
- Bbb

任务要求:

1.在文本开始和结束处,分别插入一个”@”符号;

2.在文本每行的开始和行尾分别插入”@”符号。

正则表达式：

^|$：表示匹配行开始或结束位置

任务1代码:

	Sub test1()
	    Dim reg, s$
	    s = "Aaa" & vbLf & "bbb"   '这里用vblf 表示行之间的换行符
	    Set reg = CreateObject("vbscript.regexp")
	    reg.Pattern = "^|$"
	    reg.Global = True
	    s = reg.Replace(s, "@")
	    MsgBox s
	End Sub

讨论:

Msgbox 最后显示的结果为:

@Aaa

Bbb@

代码中修改了global的默认属性值,设置为true；目的是保证能找到并替换全部的开始或结束位置。如果保持默认属性，则只会在开始处插入一个@号。

正则对象Reg的其它两个属性保持为默认。因为本任务无关乎字母大小问题，所以IgnoreCase属性无需要设置为Ture(当然如果设置为true,对最后结果也无影响);由于Mutiline属性保持默认,其值为False,所以整个文本只有一个开始位置和一个结束位置。

代码中使用了对象reg的replace方法,它的作用是,将在目标文本中找到的匹配（开始和结束位置）替换为”@”字符,在这里实际上是插入。然后把修改后的文本返回，重新赋值给字符变量S。

任务2代码：

	Sub test2()
	    Dim reg, s$
	    s = "Aaa" & vbLf & "bbb"
	    Set reg = CreateObject("vbscript.regexp")
	    reg.Pattern = "^|$"
	    reg.Global = True
	    reg.MultiLine = True
	    s = reg.Replace(s, "@")
	    MsgBox s
	End Sub

讨论:

任务2代码与任务1代码唯一区别是修改了mutiline默认属性,设置为True。这就意为着,该文本的每一行都存在一个开始位置和结束位置。所以Msgbox最后显示的结果为:

@Aaa@

@Baa@

6.应用对象的方法代码段

VBScirpt正则对象的方法共有三个：你可以根据任务要求选择使用一个或多个方法.

(1)TEST方法

语法:Object.Test(string)

Test方法只是简单测试目标文本中,是否包含正则表达式所描述的字符串。如果存在，则返回True,否则返回False。

例：用代码检测用户的输入是否是一个电子邮箱。

	Sub ChkEmail()
	    Dim reg, s$
	    s = InputBox("请输入一个电子邮箱:")
	    Set reg = CreateObject("vbscript.regexp")
	    reg.Pattern = "^\S+@\S+$"
	    If reg.Test(s) Then
	        MsgBox "你输入的电子邮箱格式正确:  " & s
	    Else
	        MsgBox "你输入的电子邮箱格式不正确!"
	    End If
	End Sub

讨论:

代码从用户那里获得字符串,然后赋值与字符变量S。验证邮箱的正则表达式非常简略,元字符序列"\S"表示不是空格的任意一个字符,后面紧跟一个+号表示一个以上字符。这个表达式事实上只验证了用户的输入里，在字符串之间是否有一个@符号。它甚至认为”0@中”都是正确的。下面给出一个更为严格的电子邮箱正则表达式：“^[\w.-]+@[\w.-]+$”当然要严格按电子邮箱规范写出正则表达式，可能就十分复杂，由于我们刚刚接触正则，就不在详细讨论了。

这里要关注的是，test方法的语法，在方法与正则对象之间也是用英文点号隔开，作为参数，目标字符串用英文括号包围。在这个例子中，如果Test返回的是true，表示目标文本S中找到了正则模式的匹配。则显示正确结果,否则显示错误提示。

(2)Replace方法

替换在目标文本中用正则表达式查找到的字符串。

前面例子中语句体现其语法：s=reg.replace(s,”@”)

后面括号中的参数S,代表前面代码中设置的目标文本字符串.也就是正则表达式将要作用的目标文本.”@”是用来替换的字符串参数.前面的s是Replace方法返回的结果,它是目标文本被替换后的一个副本. 如果没有找到匹配的文本，将返回与目标文本一样的一个副本.

下面继续讨论Replace方法的第二个参数:

例子中"@"是一个字面字符，要用一对双引号包围起来。第二个参数还可以是变量、表达式。如果是变量或函数则不能用双引号包围,这一点和VBA代码规则是一致的.

上一章我们知道了如果在正则表达式中使用了元字符序列()括号，那么被圆括号包围的内容会存储在特殊变量$1中。在有些编程语言中，可以直接在正则代码外使用$1变量,而VBScript中可以并只可以在Replace方法中,作为第二参数来调用。

例子：在目标文本中的数字数据后增加上单位：KG

目标文本：“他们体重分别是：张三56，李四49，王五60。”

结果文本要求: “他们体重分别是：张三56KG，李四49KG，王五60KG。”

正则表达式：(\d+)

替换文本:$1KG

	Sub testrep()
	    Dim reg, s$
	    s = "他们体重分别是：张三56，李四49，王五60。"
	    Set reg = CreateObject("vbscript.regexp")
	    reg.Pattern = "(\d+)"
	    reg.Global = True
	    s = reg.Replace(s, "$1KG")
	    MsgBox s
	End Sub

讨论:

用正则表达式(\d+),Replace方法将在目标文本中找到三个匹配,其值分别是56,49,60。并分别把每个值保存于每一个匹配对象的$1变量中。

替换文本：”$1KG”表示每一个匹配中的$1变量值与字面字符”KG”联结,组成新字符串,用来替换找到的数据字符串。

$1是一个很特殊的变量,它由美元符号与数字编号组成.如果正则表达式中有两个或两个以上的捕获性括号,则按照左括号”(“从左到右顺序编号,自动命名为$1,$2,$3….,共支持99组.要指出的是,如果找到多个匹配,那么每个匹配中的特殊变量名是一样的.这个例中共有三个匹配其值分别为56,49,60.第一个匹配的变量名是$1,第二和第三个匹配的变量名仍然是$1,只是每个匹配中$1保存的值是不一样的.

最后一点,作为替换参数的一部分,$1变量与字面字符共同组成替换字符串时,它们之间不用 & 符号连接,并且 $1 必须放在一个双引号中;而如果是用其它普通变量与字面字符联结组成替换文本时,则必须用 & 符号联接,这一点与VBA代码使用方法相同.

在Replace方法的第二个参数中,还有几个很少用到的特殊变量:

	$*或$&	匹配到的字符串
	$`	匹配字符串之前的文本
	$'	匹配字符串之后的文本
	$_	目标文本

一个较特殊的状况,如果上面所述的特殊变量符不是作为变量使用,而是要以它们作为字面字符的替换文本,那么就要对它们转义,方法是在它们之前加一个美元符号$.如$$&

(3)Execute方法

在目标文本中执行正则表达式搜索。

语法:set mh=object.execute(s)

其中mh是用户自定的对象变量,S是值为目标文本的字符串变量.object是正则对象.

Execute方法会作用于目标文本(S),并返回一个叫作"Matches"的集合对象,在这里是mh.在这个集合对象中包含它找到的所有叫做"Match"的成功匹配对象(Matches集合最多可容纳65536个匹配对象). 如果未找到匹配，Execute 将返回空的 Matches 集合。Matches集合有两个只读属性:索引(Item)和成功匹配的次数(Count).

Matches集合中包含的匹配对象Match有四个只读属性:Value/firstindex/length/submatches

值得一提的是,Submatches属性是一个集合属性,集合中元素个数与正则表达式中使用的捕获性括号的个数相同,每个元素的值就是括号包围起来的内容.它也有两个只读属性:item和Count
下面用树状图来表示它们之间的关系,并在接下来的内容中继续逐一讨论它们的用法.


<1>Matches集合的Item和Count属性

利用Matches集合的Item属性可以得到它包含的每个Match对象;利用Count属性可以得到成功匹配的个数.

Matches集合对象中元素(成功匹配)的索引编号从0开始.我们可以用遍历集合的方式或索引方法读取每一个匹配值.

例:从一段文本中提取所有英文单词.

目标文本:”苹果:iphone_5s;诺基亚:Nokia_1020”

结果要求:分别提取出iphone_5s和Nokia_1020

\w+	\w是元字符序列，表示英文单词字符（a-z,A-Z,0-9,_)，后面紧跟这元字符"_",表示匹配连续存在的一个或多个英文单词字符

代码:
	
	Sub test2()
	    Dim reg, k, mh, strA$
	    strA = "苹果:iphone_5s;诺基亚:Nokia_1020"
	    Set reg =CreateObject("vbscript.regexp")
	    reg.Pattern = "\w+"
	    reg.Global = True
	    Set mh = reg.Execute(strA)
	    For Each mhk In mh
	        Debug.Print mhk.value
	    Next
	End Sub

讨论:

通过语句Set mh = reg.Execute(strA),Execute方法返回一个集合对象mh,在这个集合对象里包含两个匹配对象,代码中用遍历方法取出每一个匹配对象的值.

Execute方法返回的集合对象mh,有两个属性:

1)Count:Execute方法成功匹配的次数,也可理解为mh集合对象中包含的成功匹配对象的个数.语法:

	N=mh.count   本例中n值为2

2)Item: 索引,可以通过索引值,返回集合对象中指定的匹配对象.语法:

	Set mhk=mh.item(0) 
	K=mhk.value

用索引返回第一个Match对象即mhk. 本例中k为第一个Match对象的值(iphone_5s). 同样的方法可以得到第二匹配的值.

由于Item和Value属性是集合的默认属性,所以上面两个语句也可简写为:

K=mh(0)......第一个匹配对象的值(iphone_5s)

M=mh(1)...........第二个匹配对象的值(Nokia_1020)

上面代码中遍历集合也可以用索引法遍历:

	For i=0 to mh.count-1
	    Debug.print mh(i).value
	Next i

<2>Match对象的属性

Execute方法返回的集合对象中包含的也是对象元素,即match对象,match对象有四个属性:

- FirstIndex：匹配对象所匹配字符串的起始位置。
- Length：匹配对象所匹配字符串的字符长度。
- SubMatches：匹配对象所匹配结果中的子项集合。
- Value：匹配对象所匹配的值。

在本例中:索引为0,即第一个匹配对象的属性值为:

K=mh(0).value   k的值为iphone_5s,value是默认属性可简写为k=mh(0)

sn=Mh(0).firsindex   sn的值为3,表示在目标字符串中,位置3上找到该匹配iphone_5s.(位置是从0开始的)

Ln=mh(0).length   ln值为9,即iphone_5s的字符长度

<3>Match对象的Submatches属性  

匹配对象match的Submatches是一个集合属性,它包含正则表达式中用圆括号捕捉到的所有子匹配.它为用户提供了返回$1特殊变量值的方法.

集合Submatches有两个固有属性:Count和Item.可以通过Item得到集合中的每个值,它实际就是在正则表达式中用圆括号捕获的内容;Count值是集合中元素个数,实际上就是正则表达式中捕获性圆括号的个数.

下面给一个实例来说明:

目标文本:给定一个标准邮箱地址:J3721@163.com

要求:从邮箱中分别提取出:用户名j3721,服务器域名163.com

正则表达式:  ^(\w+)@(.+)$

代码:

	Sub test5()
	     Dim reg, mh, strA$, username$, domname$
	     strA = "J3721@163.com"
	     Set reg = CreateObject("vbscript.regexp")
	     reg.Pattern = "^(\w+)@(.+)$"
	     Set mh = reg.Execute(strA)
	     N=mh(0).submatches.count         ‘n值等于2
	     username = mh(0).submatches(0)    ‘j3721
	     domname = mh(0).submatches(1)    ‘163.com
	End Sub

讨论:

正则表达式中,\w+表示匹配@前面的所有英文单词字符;@后面的点号是一个元字符,表示匹配除换行符外的所有字符之一,后面紧跟+号,即”.+”表示匹配@后面除了换行符外的所有字符.用括号包围起来,用户名和域名就会自动分别保存在变量$1和$2中.

前面已经知道VBA不能在replace之外直接调用$1或$2,而这个例子告诉我们可以用match对象的submatches集合属性来提取.

在这个例子中,execute方法返回的集合对象mh中,mh中只有一个匹配对象Match,即mh(0);mh(0)对象的属性submatches(0),返回第一个括号中的内容,即j3721.而submatches(1),返回第二个括号中的内容.submathches集合也有count属性,所以如果有很多子项需要提取,也可用遍历或索引方法返回每一个特殊变量值.最后再给一例子:

下面的代码演示了如何从一个正则表达式获得一个 SubMatches 集合以及它的专有成员：
正则表达式(一个邮箱地址):

	(\w+)@(\w+)\.(\w+)

如果你没有进一步了解元字符,可能不懂其中含义,不过没关系,在这里你只要知道,该代码的任务是显示电子邮箱dragon@xyzzy.com,用户名和组织名.
	
	Function SubMatchTest(inpStr)
	  Dim oRe, oMatch, oMatches
	  Set oRe = New RegExp
	  ' 查找一个电子邮件地址
	  oRe.Pattern = "(\w+)@(\w+)\.(\w+)"
	  ' 得到 Matches 集合
	  Set oMatches = oRe.Execute(inpStr)
	  ' 得到 Matches 集合中的第一项
	  Set oMatch = oMatches(0)
	  ' 创建结果字符串。
	  ' Match 对象是完整匹配 — dragon@xyzzy.com
	  retStr = "电子邮件地址是： " & oMatch & vbNewline
	  ' 得到地址的子匹配部分。
	  retStr = retStr & "电子邮件别名是： " & oMatch.SubMatches(0)   ' dragon
	  retStr = retStr & vbNewline
	  retStr = retStr & "组织是： " & oMatch. SubMatches(1)          ' xyzzy
	  SubMatchTest = retStr
	End Function
	Sub SubMatchesTest()
	   MsgBox(SubMatchTest("请写信到 dragon@xyzzy.com 。 谢谢！"))
	End Sub

如果知道一点英文，理解记忆会更快。

\w→ word 首字母 w 

表示26个英文字符【A-Za-z】以及下划线【_】和数字【0-9】的集合，

.Pattern ="\w" 等价于 .Pattern ="[0-9a-z_A-Z]"

其中有英文字符很好理解，因为英语中的word肯定是由英文字母构成的。

另外，在VBA编程中，变量还可以含有数字和下划线，因此数字和下划线也被当做构成统一word的要素。

除此之外的其它字符，都不算构成word的要素了。

----------

\d→ digit 首字母 d

表示数字【0-9】的集合，

.Pattern ="\d" 等价于 .Pattern ="[0-9]"

\s→ separate 首字母 s   或space、tab、return简称str字符的首字母。

表示分隔符号 含space 空格【 】或【char(32)】、回车vbCr【char(13)】、换行vbLf【char(10)】、vbTab【char(9)】等，

全部ASCII码值为： char(9)、char(10)、char(11)、char(12)、char(13)、char(32)（空格）

分隔符号的单独分开是\t   tab 首字母  vbTab【chr(9)】或 其反集 \T  

\v   verticaltab 首字母  vbVerticalTab【chr(11)】或 其反集 \V  

\f   formfeed 首字母  vbFormFeed【chr(12)】或 其反集 \F

\r   回车 return 首字母 r 而在VBA中对应的是：vbCr【chr(13)】或 其反集 \R   （Cr是Carriage Return的简称，是机械式打字机时代，字车carriage回复return 到最左边开始的意思 ）

苹果机(MAC OS系统) 采用回车符Cr 表示下一行.

\n   换行 newline 首字母n  vbLf【chr(10)】或 其反集 \N 
（Lf是Line Feed的简称，是机械式打字机时代，回车的同时自动滚进一行的意思 ）
UNIX/Linux系统采用 换行符Lf 表示下一行.

\r\n 回车换行  vbCrLf【chr(13) & chr(10)】
Dos和windows系统 采用 回车+换行CrLf 表示下一行,
----------

接下来介绍，对上述几个常用元序列/集合，使用小写字母时是有效，含有的意思，
而使用大写字母时是无效，排除（不含有）的意思，

如：

1. \W → word 首字母 大写 W 标记的【英文字母、数字、下划线】的序列/集合的反集

小写字母w .Pattern ="\w" 等价于 .Pattern ="[0-9a-z_A-Z]"

大写字母W .Pattern ="\W" 等价于 .Pattern ="[^0-9a-z_A-Z]"

2. \D → digit 首字母 大写 D 标记的【数字】的序列/集合的反集
小写字母d .Pattern ="\d" 等价于 .Pattern ="[0-9]"
大写字母D .Pattern ="\D" 等价于 .Pattern ="[^0-9]"

3. \S → separate 首字母 大写 S 标记的不可见换行字符的序列/集合的反集
小写字母s .Pattern ="\s" 等价于 含有char 9,10,11,12,13,32
大写字母S .Pattern ="\S" 等价于 不含有char 9,10,11,12,13,32
\b 和 \B begin 的首字母 b

意义： 以上述\s separate 集合分隔后得到的每一个 \w word 中的第1个begin字符。

如：

abe sau

dty12 f_34

执行小写字母b .Pattern ="\b." 则得到每一个word的首字母：a s d f 

执行大写字母B .Pattern ="\B." 则得到每一个word的首字母以外的字符：b e a u t y 1 2 _ 3 4


补充，该规则仅对英文word能100%正确作用。

如果对象字符集中出现\W字符即非英文数字下划线字符，如汉字或其他符号，
则可能会返回不可预知的结果。（因为汉字等其它词汇无法像英文那样简单区分word和word之间的间隔。）

（如同Vlookup函数中搜寻数组中没有正确A-Z排序时可能返回不可预知结果一样。）

(前面第一篇的第一/二部分,概要阐述了正则表达式的基本思想,并对正则在VBA中的实现(也就是Regexp对象操作)作了详细讲解.接下来的第二至第六部分我们将集中介绍VBA中(本质上是Regexp对象中)可使用的全部元字符(序列),有的称为元字符的"特性",有的叫作正则"语法".反正它的基本属性就是用来描述字符的特殊字符.)

