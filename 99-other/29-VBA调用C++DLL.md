[http://club.excelhome.net/forum.php?mod=viewthread&tid=1142031&extra=page%3D1&page=1&](http://club.excelhome.net/forum.php?mod=viewthread&tid=1142031&extra=page%3D1&page=1&)

# VBA调用C++的dll简易指南 #

首先，要在Excel中调用dll的函数……

C++（注意是C++不是C）端要这样写：Visual Studio版

    extern "C" __declspec(dllexport) double __stdcall 
	{
	    return *a + b;
	}

GCC（MinGW）版

	extern "C" double Add(double *a, double b) __attribute__((stdcall));
	
	double Add(double *a, double b)
	{
	return *a + b;
	}

VBA端：

	Declare PtrSafe Function Add Lib "DLLTest.dll" (ByRef x As Double, ByVal y As Double) As Double

注意以上下划线都是两个。可以看出，C++端的函数声明多了3个内容：

1 extern "C"：表明函数名按照C的习惯编译。因为C++比C多了个重载，库中的函数名可能会被编译器添上各种修饰尾巴，用这个可以禁掉那些尾巴

2 declspec(dllexport)：表明这个函数被摆上了货架子可以被外部的程序调用。为啥只有VS有这个？因为GCC会自动把所有extern的都摆上去……

3 stdcall：大概就是参数清栈规则一类滴东东。默认的cdecl是调用者清栈，实际上VBA可不会去给C++擦屁股，所以加上stdcall让被调者自己清栈

VBA端的写法很简单，查下Declare的说明就知道了。注意那个PtrSafe，是Excel 2007没有、2010开始32位可加、64位必须加的。关于这个时代32位和64位共存带来的种种麻烦，后面会单独讨论。

会写函数形式后，下一个问题就是怎么传参数，怎么传字符串，怎么传数组……其实这里的对应原则很简单：ByVal传值，ByRef传指针。变量类型的对应关系如下：

	VBA		C++				字节数
	Long	int				4
	double	double			8
	string	char* （定长）	1×长度

这三个是比较保险的，其它类型的就不太好讲了……值得注意的是，VBA的string在传给C++后会变成定长的char*，也就是说长度信息会一起传过去。建议一接到后立刻将其用C++自己的string存起来：

	#include <string>

至于如何传数组或复杂的数据结构……建议只要东东一多就上结构体吧，像这样：
C++端：

	struct ABC{
		int a;
		double b;
		int c[12];
	};

VBA端：

	Type ABC
		a as Long
		b as Double
		c(0 to 11) as Long
	End Type

调用函数时直接用ByRef/指针把结构体地址传过去~注意两边要保持一致哦。另外如果在32位系统上试验这个结构体会出现错误，下面会详细分析~


还有dll文件放置的地方也是个问题……如果要放在同一目录下，建议最好在一开始就把当前目录拽过来：

	Sub Workbook_Open()
		ChDrive (ThisWorkbook.Path)

## 32位和64位带来的困扰 ##

目前这个时期，32位和64位共存带来了许多困扰，如果要发布东东的话就得两边都照顾到。在VBA和C++联用的环境下，这里有好几个小陷阱……

首先要明确几个问题：

1 说起【32位】和【64位】，指的可能是处理器、系统或是程序，这里是单指Office的位数，因为64位Windows也可能装了一个32位的Office。

2 32位和64位到底有多大区别？有很多传言，说是int变64位啦之类的。实际上int是4字节、double是8字节这两个长度已经相当深入人心了，随便乱动的话会引发动乱的。比较确定变化的是指针的位数，而其它的就……比如指针的真身，C++的Long和unsigned Long就有可能跟着升到了8字节64位……

3 同一个结构体，就算只由char、int和double组成，在32位和64位中的储存方式也可能不一样，这是引发混乱的根源之一……

在VBA调用C++dll的场合关于位数，有3个问题要注意：

### 1 位数对应 ###

dll的位数必须和Office的位数一致。dll的位数取决于编译器，和自己写的代码没啥关系，如果用VS的话就要在解决方案的配置属性里设置下平台（Win32还是x64），GCC的话MinGW到64位就不灵了，得去下个MinGW w64，这个东东有32位和64位的版本。至于判断Office的位数有个笨办法，去Program Files和Program Files (x86)下的Office目录里看下哪边东西多……

如果位数不对，VBA运行时就会提示找不到那个库文件，不管把库放在哪里都一样

### 2 函数名尾巴 ###

目前64位的编译器编出来的函数名没发现啥问题，32位的编译器编出的dll函数名后面还是经常会带着一个尾巴，似乎是@+参数总位数的样子。对付这一点最简单的办法就是，先去下个Depends搞清楚编出来的函数到底叫什么，然后在VBA的声明里分开写：

	#If Win64 Then
		Declare PtrSafe Function Add Lib "DLLTest.dll" (ByRef x As Double, ByVal y As Double) As Double
	#Else
		Declare Function Add Lib "DLLTest.dll" Alias "Add@12" (ByRef x As Double, ByVal y As Double) As Doubler
	#End If

注意，VBA的预编译宏Win64返回的刚好是Office是不是64位，而非Windows是否64位。

### 3 讨厌的结构体 ###
结构体在内存中的存储方式可不是一个挨一个的，有一个叫做【对齐】的很复杂的问题，这里就不具体分析了，感兴趣可以问度受 。在64位平台上，如果只用int、double和string的话一般不会有事，但32位平台上……据大致分析，MinGW w64不论64位版还是32位版对齐值都是8，64位VBA的对齐值也是8，但32位VBA的对齐值是4。回顾下上面说有问题的那个结构体：

	Type ABC
		a as Long
		b as Double
		c(0 to 11) as Long
	End Type

在64位平台下，每个变量都占着8个字节的屋子，a只有4个字节的话后面4个字节就是空的。32位平台的MinGW w64也是这样干的，不知道VS是不是这样干，但VBA中对齐值变成了4，a只能住刚好标配的4字节屋子，紧跟着b就挤进来了，于是在32位VBA和C++（MinGW w64）中两种结构体里b的位置就差了4字节，传数据时当然就乱了套。

解决这个问题的办法，一是写结构体时注意变量顺序，一定要让4字节的变量两两凑成对，但这样很麻烦，尤其是遇到数组时经常算不清。这里建议从C++那边想办法，用预编译宏：

	#ifndef _WIN64
	#pragma pack(4)
	#endif

让编译器在32位系统中强制规定对齐值为4。