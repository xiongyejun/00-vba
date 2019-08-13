'15位的身份证编码首先把出生年扩展为4位，简单的就是增加一个19，但是这对于1900年出生的人不适用
'6位-地址码
'8位-出生日期码
'3位-顺序码：顺序码的奇数分给男性，偶数分给女性
'1位-校验码
'校验码是根据前面十七位数字码，按照ISO 7064:1983.MOD 11-2校验码计算出来的检验码
'   1、将前面的身份证号码17位数分别乘以不同的系数。从第一位到第十七位的系数分别为：7 9 10 5 8 4 2 1 6 3 7 9 10 5 8 4 2 ；
'   2、将这17位数字和系数相乘的结果相加；
'   3、用加出来和除以11，看余数是多少；
'   4、余数只可能有0 1 2 3 4 5 6 7 8 9 10这11个数字。其分别对应的最后一位身份证的号码为1 0 X 9 8 7 6 5 4 3 2
Function CheckID(ID As String) As Boolean
    Dim i As Long
    Dim t As Long
    Dim s As String

    For i = 1 To 17
        t = t + VBA.CLng(Mid(ID, i, 1)) * (2 ^ (18 - i) Mod 11)
    Next
    s = Mid("10X98765432", t Mod 11 + 1, 1)

    If s = Right(ID, 1) Then
        CheckID = True
    Else
        CheckID = False
    End If
End Function

Sub test()
    Debug.Print CheckID("231182194012198424")
End Sub
