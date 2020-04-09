# 数据类型Array

标签（空格分隔）： 数据类型

---

VBA里的数据数组Array，底层结构是SafeArray，[官方文档][1]。
获取数组地址的API：

    Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Var() As Any) As Long

测试一下这个API

    Type SafeArrayBound
      cElements As Long '该维的长度
      lLbound As Long '该维的数组存取的下限，一般为0
    End Type
    
    Type SafeArray
        cDims As Integer ' 数组的维度
        fFeatures As Integer
        cbElements As Long ' 数组元素的字节大小
        cLocksas As Long
        pvDataas As Long '数组的数据
       rgsabound(1) As SafeArrayBound
    End Type
    
    Sub TestArray()
        Dim b() As Byte
        Dim bptr As Long
        
        ReDim b(3) As Byte
        bptr = VarPtrArray(b)
        
        Dim sa As SafeArray
        Copymemory VarPtr(sa.cDims), bptr, Len(sa)
        
        Printf "bptr = 0x%x, sa.cDims = %d", bptr, sa.cDims
    End Sub
    输出：
    bptr = 0x2eef50, sa.cDims = 14296
    
sa.cDims应该返回1才对，这里明显是不对的，所以猜测VarPtrArray返回的还不是SafeArray结构的地址，像C语言，一般在声明结构的时候，是用指针的，所以这个很有可能是1个指针，指向SafeArray结构。

    Sub TestArray()
        Dim b() As Byte
        Dim bptr As Long
        
        ReDim b(3) As Byte
        bptr = VarPtrArray(b)
        
        Dim bptrptr As Long
        Copymemory VarPtr(bptrptr), bptr, 4
        
        Dim sa As SafeArray
        Copymemory VarPtr(sa.cDims), bptrptr, Len(sa)
        
        Printf "bptr = 0x%x, bptrptr = 0x%x,sa.cDims = %d, sa.cbElements = %d, sa.rgsabound(0).cElements = %d, sa.pvDataas = 0x%x, VarPtr(b(0)) = 0x%x", bptr, bptrptr, sa.cDims, sa.cbElements, sa.rgsabound(0).cElements, sa.pvDataas, VarPtr(b(0))
    End Sub
    输出：
    bptr = 0x2eef50, bptrptr = 0x23834be8,sa.cDims = 1, sa.cbElements = 1, sa.rgsabound(0).cElements = 4, sa.pvDataas = 0x13061798, VarPtr(b(0)) = 0x13061798



  [1]: https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-safearray