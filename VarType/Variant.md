# 数据类型Variant

标签（空格分隔）： 数据类型

---

[官方文档定义][1]：

    A special data type that can contain numeric, string, or date data as well as user-defined types and the special values Empty and Null. The Variant data type has a numeric storage size of 16 bytes and can contain data up to the range of a Decimal, or a character storage size of 22 bytes (plus string length), and can store any character text. The VarType function defines how the data in a Variant is treated. All variables become Variant data types if not explicitly declared as some other data type.

16字节是如何分配的？

 1. 前8字节
    b0：标识数据类型（[官方文档][2]有详细定义）
    b1：标识后8字节是数据还是指针
		- 0x00	数据类型的是数据本身，String是地址
		- 0x40	8-11存的是数据地址，String是地址的地址
		- 0x20	8-11存的是数组地址
		- 0x60	8-11存的是数组地址的地址
 2. 后8字节：数据或指针


          Sub TestVariant()
            Dim v As Variant
            Dim i As Byte
            
            i = &H10
            v = i
            
            Dim lenth As Long
            lenth = 16
            
            Dim b() As Byte
            ReDim b(lenth - 1) As Byte
            
            CopyMemory VarPtr(b(0)), VarPtr(v), lenth
            Printf "VarType(v) = 0x%x, b = 0x% x", VarType(v), b
        End Sub
        
        输出：
        i定义byte：VarType(v) = 0x11, b = 0x11 00 00 00 00 00 00 00 10 00 00 00 00 00 00 00
        i定义Integer：VarType(v) = 0x2, b = 0x02 00 00 00 00 00 00 00 10 00 00 00 00 00 00 00
        i定义Long：VarType(v) = 0x3, b = 0x03 00 00 00 00 00 00 00 10 00 00 00 00 00 00 00
        i定义Double：VarType(v) = 0x5, b = 0x05 00 00 00 00 00 00 00 00 00 00 00 00 00 30 40
        i定义String：VarType(v) = 0x8, b = 0x08 00 00 00 00 00 00 00 74 02 f3 23 00 00 00 00
        
b1一直都是0，就算用v = VarPtr(i)，仍然是0，因为VarPtr返回的也是Long，如何才能让b1出现呢？
我们知道，VBA里面，函数的传值默认就是byref，所以加1个Function就可以了

        Sub TestVariant()
            Dim i As Byte
            
            i = &H10
            TestVariantPtr i
        End Sub
        
        Function TestVariantPtr(v As Variant)
            Dim lenth As Long
            lenth = 16
            
            Dim b() As Byte
            ReDim b(lenth - 1) As Byte
            
            CopyMemory VarPtr(b(0)), VarPtr(v), lenth
            
            Dim ptr As Long
            CopyMemory VarPtr(ptr), VarPtr(b(8)), 4
            
            Dim Value As Byte
            CopyMemory VarPtr(Value), ptr, 2
            
            Printf "VarType(v) = 0x%x, b = 0x% x, ptr = 0x%x, Value = %x", VarType(v), b, ptr, Value
        End Function
        输出：
        i、Value定义byte：VarType(v) = 0x11, b = 0x11 40 00 00 00 00 00 00 52 ef 19 00 00 00 00 00, ptr = 0x19efe4, Value = 10
        i、Value定义Integer：VarType(v) = 0x2, b = 0x02 40 00 00 00 00 00 00 52 ef 19 00 00 00 00 00, ptr = 0x19efe4, Value = 10
        i、Value定义Long：VarType(v) = 0x3, b = 0x03 40 00 00 00 00 00 00 50 ef 19 00 00 00 00 00, ptr = 0x19ef50, Value = 10
        其他不多演示，注意Dim Value As语句下面的CopyMemory复制字节数和Value类型保持一致


  [1]: https://docs.microsoft.com/zh-cn/office/vba/language/glossary/vbe-glossary#variant-data-type
  [2]: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function?f1url=https://msdn.microsoft.com/query/dev11.query?appId=Dev11IDEF1&l=zh-CN&k=k%28vblr6.chm1009057%29;k%28TargetFrameworkMoniker-Office.Version=v16%29&rd=true