Attribute VB_Name = "模块2"
Option Explicit

Sub RowCnt()
    '这个小程序只用了一句话将txt文件读取到了数组中，太经典了，神一样的代码啊！
    t = Timer
    Dim Arr, k&
    fnum = FreeFile '这是一个内部赋值语句
    Open "D:\测试数据\手套.txt" For Input As #fnum
    '以下语句的解释，因为太经典了，不得不说明一下以免以后忘了：
    '首先LOF语句表示用 Open 语句打开的文件的大小，该大小以字节为单位。
    '然后InputB假定数据已是二进制，对其不加转换即存为变量
    'StrConv 函数按指定类型转换，这里是 vbUnicode
    'Split函数返回一个下标从零开始的一维数组，它包含指定数目的子字符串，分开的键值是回车 vbCrLf
    Arr = Split(StrConv(InputB(LOF(fnum), 1), vbUnicode), vbCrLf)  '按【回车键】拆分这个字符成为个数组
    k = UBound(Arr) + 1
    Reset
    Close #fnum
    MsgBox "耗时：" & Timer - t
    MsgBox "此文本总行数为" & k
    '测试一个6MB的txt文件只用了不到4秒
End Sub

