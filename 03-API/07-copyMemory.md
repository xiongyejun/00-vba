[http://club.excelhome.net/forum.php?mod=viewthread&tid=1282032&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=1282032&extra=page%3D1)

# 简单讲讲Copymemory这个函数 #

我尽量讲得简单通俗，有些说法甚至只是比喻，这就会忽略复杂性、也可能有失准确，但对初学者来说却是更容易理解，讲之前我搜了一下关于这个函数及指针的帖子，有几个贴子可谓专业翔实，非初学者可以看那些贴子，因水平所限，所讲难免有误，望大家指正，一起把问题搞清楚。

## 一、长整型 ##

	Sub Test1()
	　　　　Dim p As Long
	　　　　‘Debug.Print Hex(VarPtr(p))
	　　　　p=20008
	　　　　‘Debug.Print Hex(20008)
	End Sub

1 当执行dim p as long这一句代码时，发生了什么？

　　此时，程序会在内存中分配一个4字节的空间，用0填充，以备写入数据。

　　请看我执行此句后的内存截图：

![](http://files.c.excelhome.net/forum/201606/05/075951x2al5v1bqlky7613.png)

绿框处为分配的空间，红框处005FF1B0为空间的首字节地址，它可以用varptr函数得到。

2 当执行p=20008时，程序会把20008转成16进制数4E28，然后从低位至高位依次填入保留的空间，所以填充次序是：284E，下面的截图是当程序执行完这句时的内存状态：

![](http://files.c.excelhome.net/forum/201606/05/080517iu0l0v90ku9wsfc0.png)

## 二、字节型、整形，单精度，双精度，布尔，货币，Variant ##

当声明为上面的数据类型时，发生的情况与长整形的情况只有一点不同：

分配的预留空间长度不同，比如双精度会分配8个字节的空间，其实大家都知道，我就不说了。

## 三、字符串 ##

程序如下：

	Sub Test2()
	　　　　Dim str1 As String
	　　　　‘Debug.Print Hex(VarPtr(str1))
	　　　　str1="abcde"
	　　　　‘Debug.Print Hex(StrPtr(str1))
	End Sub

当程序执行dim str1 as string时，发生了什么？，请看内存截图：

![](http://files.c.excelhome.net/forum/201606/05/080650se6o0io30c64u4qi.png)

程序为str1在内存中分配了一个4个字节的空间，但一个字符需要2个字节的空间存储，四个字节的空间只能存2个字符啊，这是怎么回事？

事情是这样嘀：

它分配的这4个字节的空间，其实是保存一个长整形数值的，这个长整形数值是一个内存地址。

当执行str1=”abcde”时，可以看到这个空间被0E054EF4填充了：

![](http://files.c.excelhome.net/forum/201606/05/080905wtvytnido69v966t.png)

注意这个填充是按低位到高位的顺序填充的，我们说过这是一个地址，那我们就转到这个地址看一下，见图：

![](http://files.c.excelhome.net/forum/201606/05/080933z7biww6tuw4if4vo.png)

说明如下：

　　1. 在这个地址处，我们看到了字符串”abcde”，”abcde”被保存在绿框处。

　　大家注意，所有字符都是用一个2字节整形数表示的，比如a就是用&H0061表示的，你可以查一下ASCII码表，a表示为&H61，而这里&H0061是Unicode表示法。或者在立即窗口中执行?chr(&H0061)，就明白了。

　　但不要忘了，数字在内存中是从低位到高位填充的，所以a在内存中就变成6100，这样内存中61006200630064006500就表示abcde，即圈2绿框处。

　　2. 那我们如何知道字符串在哪里结束呢？你看圈3处0000（表示空字符），这其实就是用一个空字符来表示字符串结束的位置。

　　3. 但是第2条中用空字符表示结束有问题，因为我们的字符串中可能包含空字符，遇到这种情况你怎么判断字符串在哪结束？

　　4. 第3条中的问题如何解决呢？从0E054EF4这个指针（即首位地址）往前数4个字符，圈4处，我们会看到一个长整形数字：0000000A，转成十进制数即是10，这个就是字符串的长度。

　　示意图如下：

![](http://files.c.excelhome.net/forum/201606/05/081023bbe8aeyi33ewimiw.png)

简单总结一下：

　　当声明一个字符串时，会在内存中预分配一个四字节的空间，这个空间的首地址可以用varptr(str1)得到。这个空间存的不是字符串本身，当我们给字符串赋值时，程序会在内存中另辟一块空间存放字符串本身，而把另辟的那块空间的首地址写在这个预留的四字节的空间中。这个地址可以用strptr(str1)得到。

## 四、数组 ##

数组的情况又如何呢?

程序如下:
	
	Sub Test3()
	　　　　Dim arr(2)
	　　　　‘Debug.Print Hex(VarPtr(arr(0)))
	　　　　arr(0) = "大连"
	　　　　arr(1) = 900
	　　　　arr(2) = 800.5
	　　　　‘Debug.Print Hex(StrPtr(arr(0)))
	End Sub

1 当程序执行到 Dim arr(2) 时，vba会在内存中划分出一块16 * 3 大小的空间，因为数组是Variant数据类型的，有3个元素，Variant类型占16个字节，所以此数组总共需要16*3大小的空间。

见图：

![](http://files.c.excelhome.net/forum/201606/05/150956wsbd13b473sdq44n.png)

2 这个空间的第一个地址(也即此数组的指针)是0DF2E8E8，可以用varptr(arr(0))得到，可想而知，这个空间要被数组第一个元素，第二元素…依次填充，所以第一个元素的首位地址即是此空间的首位地址。

3 当程序执行完

	　　　arr(0) = "大连"	
	　　　arr(1) = 900	
	　　　arr(2) = 800.5

　　会发生什么呢？我们看下图做一说明

![](http://files.c.excelhome.net/forum/201606/05/151053u7m47b4eokk8ezbk.png)

1 先看arr(0),头2个字节存储的是数据类型，当为0008时，意味着数据类型为字符串，（各位可以查看一下vartype函数的返回值，08为字符串02为整型05为双精度）从第9位开始，数4个字节，内容为：0DF210C4，这即是内存中真正存放字符串空间的首位地址。我们转到这个地址看一下，见图：

![](http://files.c.excelhome.net/forum/201606/05/151124uzv3ozvqr6f3vuzk.png)

一处为字符串的长度，二处为字符串本身，3处为空字符，表示字符串结束。

你可以在立即窗口输入：

	　　？hex(ascb(midb("大",1)))	
	　　？hex(ascb(midb("大",2)))

　　得到27 59 

输入：

	　　？hex(ascb(midb("连",1)))	
	　　？hex(ascb(midb("连",2)))

　　得到 DE 8F

可知圈2处即为“大连”。

  2 再看arr(1)，头两个字节为00 02，这表示数据类型为Integer，这个类型占2个字节，我们从第9个字节往后数2个字节，发现是： 03 84 转成十进制即是900，但第11 12 位被填充了B3 0C，我也不明白是为什么，有知道的朋友麻烦告知一下，好在这不影响程序取数据的正确，因为已经知道是Integer类型的了，所以只取2个字节，后面的字节忽略了。

　　3. 最后看一下arr(2),这2个字节为 00 05, 表示数据类型为双精度，双度精占8个字节，我们从第9位往后数8个字节，为：4089040000000000，此16进制数转成十进制，即为800.5，进制转换各位可以搜一下百度。

　　当我们把数组声明为dim arr()时，又如何呢？其实，这种情况与前面的情况是一样的，只不过vba把分配内存空间的时刻，移到了执行redim arr(2)的时刻，你早晚都得执行这句,不是吗。

　　可见，当数组中某个元素为string类型时，vba会在预留的空间中存放一个指针(四字节)，这个指针指向真正存放字符串空间的首地址，它可以用strptr()函数得到；当元素为字节型、整形，单精度，双精度，布尔，货币数据类型时，直接存放元素的内容。

## 五、数组II ##

请看下面代码：

	　　Sub Test4()
	　　　　Dim arr(1 To 3, 1 To 2) As String
	　　　　Debug.Print Hex(VarPtr(arr(1, 1)))
	　　　　arr(1, 1) = "a"
	　　　　arr(2, 1) = "b"
	　　　　arr(3, 1) = "c"
	　　　　arr(1, 2) = "d"
	　　　　arr(2, 2) = "e"
	　　　　arr(3, 2) = "f"
	　　　　Debug.Print Hex(StrPtr(arr(3, 2)))
	　　End Sub

前面我们说过，当我们把数组声明为可变类型时，vba会自动判断每个元素的数据类型，如果是字符串类型的，则在四个字节的空间中存放一个指针（即是实际字符串存放位置的首地址），如果是整形，单精度等，则根据其长度，直接在相应位置存放其内容。

　　我们如果把数组声明为string时，会如何呢？

　　当我们执行Dim arr(1 To 3, 1 To 2) As String时，会发生什么？

　　因为此数组共3 * 2=6个元素，而已经明确是字符串型的，所以每个元素需要4个字节的空间，存放指针，所以VBA会预先在内存中划分一块4*6=24字节的空间。见下图：

![](http://files.c.excelhome.net/forum/201606/06/101541zsdpos7878fh3235.png)

6个红框的内容分别是指向真正存放字符串地址的指针。

下面我们再来看看多维数组是按什么顺序在内存中存储的，执行下面的程序：

	　　Sub Test5()
	　　　　Dim arr(1 To 3, 1 To 2) As Integer
	　　　　Debug.Print Hex(VarPtr(arr(1, 1)))
	　　　　arr(1, 1) = 1
	　　　　arr(2, 1) = 3
	　　　　arr(3, 1) = 5
	　　　　arr(1, 2) = 7
	　　　　arr(2, 2) = 8
	　　　　arr(3, 2) = 9
	　　End Sub

执行完程序的内存截图：

![](http://files.c.excelhome.net/forum/201606/06/101540duucafxdacuxobf6.png)

大家看6个红框，按顺序是1 3 5 7 8 9，可见多维数组是按列的顺序填充内存的，这一点很重要，不明白这一点，就不能正确复制数组。

　　另外，说个题外的，当你用redim 重新定义数组大小时，程序会在另一个区域新辟一块内存，供数组使用，原来数组所在的内存会被释放，而如果你要保留原来的数据（redim Preserve），程序会把原来的数据复制到新区域中。尽管这个过程非常之快，但如果redim preserve 处于一个很大的循环中时，还是会托慢程序。如何处理，因无关这个题目就不多说了。

## 六、Copymemory函数 ##

　　1. 我搜了一下讲解这个函数的帖子，发现对这个函数的声明大体有两种：

	A: Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
	
	B: Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long,ByVal Length As Long)  

　　虽然使用B有其理由，但为避免混乱，这里使用A的声明方式，初学者只需把A声明拷贝到模块开头就可以使用Copymemory函数了，其他不必深究。

　　2. 其实这个函数本身非常简单:

　　Copymemory 参数2, 参数1, a

　　记住：参数1 ，参数2 都是内存一块区域开头的地址，这个函数的意思是：把从参数1开头的地址，往后数a长度这么大的区域，复制到参数2开头的区域中。

　　3. 见下图：这个函数无非是把0015F7C0（参数1）开头的区域，红框1处的内容，复制到以0015F7C4（参数2）开头的区域，即红框2处， 这里第3个参数为4，意即复制四个字节的内容。

复制前

![](http://files.c.excelhome.net/forum/201606/07/133659xhgaoj88ozadgaat.jpg)

复制后

![](http://files.c.excelhome.net/forum/201606/07/133702s0koybyyvy8bbnyo.jpg)

大家需要明白的两点是:

　　①既然参数1,参数2是内存的地址，而内存的地址是用一个长整形数表示的，所以，参数1,参数2都是long数据类型的。

　　②这个函数的2个参数，参数1、参数2必须传址（byref），不能传值(byval)，除非这个值本身是一个内存地址。
　

　　到此为止，大家可以看到，这个函数本身是很简单的，也容易理解。

问题在于如何取得首位地址，待续……

## 七、Copymemory II ##

实在倒不出时间，做了张图，先贴在这里，大家先看看这张图，明白了这张图，也就基本明白了这个函数。

![](http://files.c.excelhome.net/forum/201606/08/095650gnmmtbpwy6dzld2d.png)


　　1. 我先截取一块示意图，给大家讲一下示意图是什么意思：

　　图A

![](http://files.c.excelhome.net/forum/201606/10/180541u703b3s33bujss80.png)

图A完美解释了 s1=”abcde”是什么意思。

　　当程序执行到dim s1 as string时，会在内存中划分一个四个字节大小的空间，这个空间第一个字节的地址如图是1000，我们以比喻的方式说，1000即是赵家的地址，或门牌号，此时赵家空着。

　　当程序执行到s1=”abcde”时，程序会在内存中划分另一块区域，这个区域多大，我们后说，这个区域第一个字节的地址是1500，并且程序会把这个门牌号放进赵家(也就是byval s1的位置)。我们把1500看做是郑家的门牌号。

　　也说是说我们按地址1000，进入了赵家，看到的不是abcde，看到的却是一个门牌号1500（byval s1），我们按这个门牌号指引去了郑家，才会看到abcde。

　　2. copymemory 参数2,参数1,a

　　我们可以把copymemory相像成一辆小车，参数1是告诉这辆车去谁家拿东西，参数2是告诉小车把东西拿到谁家，a是告诉拿多少东西。

　　我们看截图：

　　图B

![](http://files.c.excelhome.net/forum/201606/10/180541sv8u8b1avem30j0a.png)

如果我们要把s1的abcde复制到s2中，即是把郑家的abcde（圈1），放到王家（圈2）。

　　郑家的地址是1500（byval s1）,王家的地址是2500（byval s2）,所以我们要给copymemory这辆小车这样下命令：

　　Copymemory byval s2,byval s1,5

　　完整程序如下：

	　　Sub testn1()
	　　　　Dim s1 As String
	　　　　Dim s2 As String
	　　　　s1 = "abcde"
	　　　　s2 = String$(5, 0)
	　　　　CopyMemory ByVal s2, ByVal s1, 5
	　　　　Debug.Print s2
	　　End Sub

　“abcde”这个字符串，在内存中应该是占了10个字节的内容啊（61 00 62 00 63 00 64 00 65 00）

![](http://files.c.excelhome.net/forum/201606/10/180542zz6eiszgpuyk8uu9.png)

当vba发现你要传递一个字符串参数给一个外部函数（copymemory）时，它会叫暂停，然后：

　　a. 首先它把字符串由Unicode码转为ansi码

　　b. 其次，它把转成ansi码的字符串变量传给函数

　　c. 函数执行完后，它把得到的结果再由ansi码转回Unicode

　　具体见下图：

![](http://files.c.excelhome.net/forum/201606/11/204304uywwxsjj3xifd3xi.png)

关于unicode与ansi，你须知道下面一条就可以了：

　　在unicode中，中文和英文字符都占2个字节，在ansi中，中文占2个字节，英文占1个字节。

　　我们再看下面的例子：
	
	　　Sub testn1()
	　　　　Dim s1 As String
	　　　　Dim s2 As String
	　　　　s1 = "abcde混沌"
	　　　　s2 = String$(7, 0)
	　　　　CopyMemory ByVal s2, ByVal s1,7
	　　　　Debug.Print s2
	　　End Sub

虽然"abcde混沌"的长度为7，但得到的结果s2是：“abcde混”，刚才说了，转成ansi码时，英文字符占一个字节，中文字符占2个字节，这里一共有5个英文字符，2个中文字符，所以ansi的长度应该是9，把程序改为：

	　　Sub testn1()
	　　　　Dim s1 As String
	　　　　Dim s2 As String
	　　　　s1 = "abcde混沌"
	　　　　s2 = String$(9, 0)
	　　　　CopyMemory ByVal s2, ByVal s1,9
	　　　　Debug.Print s2
	　　End Sub

你会发现结果正确了。

　　那么，难道字符串的长度每回我都要数一下吗？我们如何取得这个长度呢？你可以用：

　　LenB(VBA.StrConv(s1, vbFromUnicode))

　　来取字符串长度，程序修改如下：

	　　Sub testn2()
	　　　　Dim s1 As String
	　　　　Dim s2 As String
	　　　　Dim p As Long
	　　　　s1 = "abcde混沌"
	　　　　p = LenB(VBA.StrConv(s1, vbFromUnicode))
	　　　　s2 = String$(p, 0)
	　　　　CopyMemory ByVal s2, ByVal s1, p
	　　　　Debug.Print s2
	　　End Sub

看了以上的讲解，你也许觉得有些繁琐了，又要转换，又要计算长度，有没有绕过UA转换的方法呢？当然是有，我们看下图

![](http://files.c.excelhome.net/forum/201606/11/204307hborhqvu6bh46wo2.png)

Copymemory byval s2,byval s1,9 这种方式，是通过线路1,线路2,分别到郑家和王家的，我们也可以通过线路3到郑家，通过线路4到王家，代码为：

	Copymemory byval strptr(s2),byval strptr(s1),14

　　当s1,s2在括号中时，vba就不会发现我们是在传字符串给函数，也就不会做转换了，此时，我们直接到郑家取14个长度的内容，放到王家，所以第三个参数为14，程序如下：

	　　Sub testb1()
	　　　　Dim s1 As String
	　　　　Dim s2 As String
	　　　　s1 = "abcde混沌"
	　　　　s2 = String$(Len(s1), 0)               ‘这句要用len，实际会在内存分配14个字节空间
	　　　　CopyMemory ByVal StrPtr(s2), ByVal StrPtr(s1), LenB(s1)      ‘这句要用lenb，要复制14个字节
	　　　　Debug.Print s2
	　　End Sub

我们看示意图，既然线路1,3可以到郑家，线路2,4可以到王家，那么我们可不可以这样组合：

	CopyMemory ByVal StrPtr(s2), ByVal s1，a

　　不要这样做，因为这里参数2没转码，参数1转码，所以会得出错误的结果。

