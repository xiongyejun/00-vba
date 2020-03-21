# pointer

标签（空格分隔）： 数据类型

---

## VarPtr ##
定义1个变量，函数返回这个变量的地址。

这个变量可以是任何类型。

## StrPtr ##
定义1个String类型

 - 初始化前，函数返回0，这个时候还没有字符的内存地址，所以是0
 - 初始化后，函数返回字符所在的内存地址（假设是ps）

与VarPtr得到的变量地址（假设是pv）关系是，pv这个地址保存的4个字节（32位）的值就是ps

    Sub TestString()
        Dim str As String
        Dim lVarPtr As Long
            
        str = "a"
        CopyMemory VarPtr(lVarPtr), VarPtr(str), 4
        printf "VarPtr(str) = 0x%x, StrPtr(str) = 0x%x, lVarPtr = 0x%x", VarPtr(str), StrPtr(str), lVarPtr
    End Sub
    输出：
    VarPtr(str) = 0x1ef030, StrPtr(str) = 0x17a455ec, lVarPtr = 0x17a455ec
    
    说明：
    API CopyMemory 声明
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
    
    printf 是自己封装的1个函数，代码没有列出
    
所以从上面可以看出，StrPtr就是把VarPtr得到的pv这个指针保存的内存数据读取出来。

也就是对String类型来说，其实有VarPtr就能够间接获取Str的字符所在地址，于是就想到尝试用StrPtr对数据类型的变量使用是不是也是一样的效果？

    Sub TestStrPtr()
        Dim l As Long
        
        printf "VarPtr(l) = 0x%x, StrPtr(l) = 0x%x", VarPtr(l), StrPtr(l)
    End Sub
    
    输出：
    VarPtr(l) = 0x1ef030, StrPtr(l) = 0x17a455ec
    
本以为在变量l未初始化也就是0的时候，StrPtr(l)应该返回0才对，为什么返回了一个值？这个如果也是个内存地址的话，里面又保存了什么？

    
    Sub TestStrPtr()
        Dim l As Long
        
        printf "VarPtr(l) = 0x%x, StrPtr(l) = 0x%x", VarPtr(l), StrPtr(l)
          
        Dim b(3) As Byte
        CopyMemory VarPtr(b(0)), ByVal StrPtr(l), 4
        printf "b = 0x% x", b
    End Sub

    输出：
    VarPtr(l) = 0x1ef030, StrPtr(l) = 0x118c6a5c
    b = 0x30 00 00 00

输出的这个值还挺特殊！变量l赋值的话输出也会变化，不懂为什么！

## ObjPtr ##

这个和StrPtr好像差不多，VarPtr(obj)得到的地址po，po内存保存的数据等于ObjPtr(obj)

    Sub TestStrPtr()
        Dim Rng As Range
        Dim lVarPtr As Long
        
        Set Rng = Range("A1")
        CopyMemory VarPtr(lVarPtr), VarPtr(Rng), 4
        printf "VarPtr(Rng) = 0x%x, ObjPtr(Rng) = 0x%x, lVarPtr = 0x%x", VarPtr(Rng), ObjPtr(Rng), lVarPtr
    End Sub
    
    输出：
    VarPtr(Rng) = 0x1ef030, ObjPtr(Rng) = 0x8189280, lVarPtr = 0x8189280





