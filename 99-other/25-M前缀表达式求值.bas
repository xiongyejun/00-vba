Attribute VB_Name = "M前缀表达式求值"
Option Explicit

Sub vba_main()
    Dim str As String
    Dim p_str As Integer

    str = "* + 7 * * 4 6 + 8 9 5"
    p_str = 1
   
    Debug.Print Eval(str, p_str)
   
End Sub

'前缀表达式求值的递归程序
Function Eval(str As String, p_str As Integer) As Double
    Dim x As Double
   
    x = 0#
   
    While Mid(str, p_str, 1) = " "
        p_str = p_str + 1
    Wend
   
    If Mid(str, p_str, 1) = "+" Then
        p_str = p_str + 1
        Eval = Eval(str, p_str) + Eval(str, p_str)
        Exit Function
    End If
   
    If Mid(str, p_str, 1) = "*" Then
        p_str = p_str + 1
        Eval = Eval(str, p_str) * Eval(str, p_str)
        Exit Function
    End If
   
    While Mid(str, p_str, 1) >= "0" And Mid(str, p_str, 1) <= "9"
        x = 10# * x + VBA.Val(Mid(str, p_str, 1))
        p_str = p_str + 1
    Wend
   
    Eval = x
End Function

