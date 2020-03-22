# 数据类型String

标签（空格分隔）： 数据类型

---

微软官方文档[data-type-summary][1]
   
    String (variable-length)	10 bytes + string length	0 to approximately 2 billion
    String (fixed-length)	Length of string	1 to approximately 65,400

## String变长 ##

10字节是如何分配的？猜测：

 1. 变量本身占用4个字节，用VarPtr可获取地址p
 2. VarPtr那个地址p保存的值，指向了字符的地址，p-4地址处保存的是长度信息，4个字节
 3. 另外2个是p-6处的00 88还是字符结尾的00 00？(测试将00 88修改为其他的时候,End Sub之后Excel崩溃，修改结尾的00 00则不会)
 
        
        Sub TestString()
            Dim str As String
            
            str = "a"
            
            '10 官方定义str长度
            '4  变量占用
            '2  字符长度
            '2  字符后面00 00
            Dim b(10 - 4 + 2 + 2) As Byte
            CopyMemory VarPtr(b(0)), StrPtr(str) - 6, 10
            printf "b = % x", b
        End Sub
    
        输出：
        b = 00 88 02 00 00 00 61 00 00 00 00
    

### 自己构建1个内存区域赋值给VarPtr那个地址p ###
用过一些API后，会发现很多需要返回字符串的API都是要先在VBA里先声明1个String，并且赋值一个足够的长度，调用之后再根据返回长度来取出需要的字符串。

如果API函数直接返回字符串内存指针及字符长度，在VBA里用CopyMemory也可以，但是这样就不能释放API函数传出来的内存。

我就想如果是在API里传出StrPtr需要的那个地址，赋值给1个str的VarPtr那个地址，是不是程序退出的时候VBA的垃圾回收能释放那个内存？

C代码

    __declspec(dllexport) char* __stdcall RetStrPtr()
    {
    	
    	char* ch = (char*)malloc(9);
    
    	ch[0] = 0x00;
    	ch[1] = 0x88;
    	ch[2] = 0x02;
    	ch[3] = 0x00;
    	ch[4] = 0x00;
    	ch[5] = 0x00;
    	ch[6] = 0x61;
    	ch[7] = 0x00;
    	ch[8] = 0x00;
    	ch[9] = 0x00;
    	
    	return ch;
    }
    
    编译：
    cl -c 1.c 1.def
    link -DLL -out:cdlltest.dll 1.obj

VBA调用

    Public Declare Function RetStrPtr Lib "cdlltest" Alias "_RetStrPtr@0" () As Long

    Sub TestCRet()
        Dim hdll As Long
        
        hdll = LoadLibrary(ThisWorkbook.Path & "\cdll\cdlltest.dll")
        printf "hdll = 0x%x", hdll
        
        Dim str As String
        Dim lStrPtr As Long
        lStrPtr = RetStrPtr() + 6
        CopyMemory VarPtr(str), VarPtr(lStrPtr), 4
        printf "str = %s", str
        
        Stop
        FreeLibrary hdll
    End Sub
    
    输出：
    hdll = 0x69b50000
    str = a
    
执行End Sub后，Excel直接崩溃。

于是尝试在VBA内部用byte数组构建后赋值到VarPtr，结果一样。
    
    Sub TestString()
        Dim str As String
        Dim b(9) As Byte
        
        b(1) = &H88
        b(2) = &H2
        b(6) = &H61
        
        Dim lStrPtr As Long
        lStrPtr = VarPtr(b(0)) + 6
        
        Printf "强制赋值VarPtr前，StrPtr(str) = 0x%x", StrPtr(str)
        '+6 StrPtr指向的字符开始的位置，不包含前面00 88和长度信息4个
        CopyMemory VarPtr(str), VarPtr(lStrPtr), 4
        Printf "强制赋值VarPtr后，StrPtr(str) = 0x%x", StrPtr(str)
        
        Printf "str = %s", str
        
        Stop
    End Sub

    输出：
    强制赋值VarPtr前，StrPtr(str) = 0x0
    强制赋值VarPtr后，StrPtr(str) = 0x1a207876
    str = a

Stop之前都正常，但是执行End Sub后，Excel直接崩溃。

这个到底是什么原因？


  [1]: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary