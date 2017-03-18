Attribute VB_Name = "M钢条切割"
Option Explicit

Sub vba_main()
    Dim i_col As Long
    Dim arr_price() As Long
    Dim i As Long
    Dim arr_num() As Long   '每个长度的切割数量
    Dim arr() As Long
    Dim n As Long
    
    '获取已知的每个长度对应的价格
    i_col = Range("A1").End(xlToRight).Column
    ReDim arr_price(i_col - 1) As Long
    For i = 2 To i_col
        arr_price(i - 1) = Cells(2, i).Value
    Next i
    
    'n行最后1列，对应每个长度的最优解价格
    '前面的列对应最优的切割数量
    '只要求解一次就可以，其他的都可以用这个
    n = Application.WorksheetFunction.Max(Range("A4:A11"))
    ReDim arr(n, i_col - 1) As Long
    
    For i = 4 To 11
        With Cells(i, 2)
            n = Cells(i, 1).Value
            .Offset(0, 10).Value = BottomUpCutRod(arr_price, n, arr)
            .Resize(1, i_col).Value = Application.WorksheetFunction.Index(arr, n + 1)
        End With
    Next i
End Sub
'arr_price    下标长度对应的价格
'arr          临时存储每个长度的最优解arr(n,0)，同时存储每个解的切割方案：即每种价格需要切割的个数
'n            钢条的长度
'arr_num      每个价格长度需要切割的个数
Function BottomUpCutRod(arr_price() As Long, n As Long, arr() As Long) As Long
    Dim i As Long, j As Long
    Dim i_max As Long
    Dim i_tmp As Long
    Dim max_price_col As Long
    
    max_price_col = UBound(arr_price)
    
    If arr(n, max_price_col) > 0 Then GoTo lable_quit
    
    '将钢条从左边切割下长度为i的一段，只对右边剩下的长度为n-i的一段继续进行切割，对左边的一段则不再进行切割
    For i = 1 To n
        i_max = -1
        
        For j = 1 To i
            If j > max_price_col Then
                i_tmp = arr(i - j, max_price_col)
            Else
                i_tmp = arr_price(j) + arr(i - j, max_price_col)
            End If
        
            If i_max < i_tmp Then
                i_max = i_tmp
                GetCutNum arr, i, j
            End If
            
        Next j
        arr(i, max_price_col) = i_max
    Next i
    
lable_quit:
    BottomUpCutRod = arr(n, max_price_col)
End Function
'记录每个价格长度的最优解切割方案
'i  当前切割的钢条的长度
'j  左边切割为j的长度
Function GetCutNum(arr() As Long, i As Long, j As Long)
    Dim k As Long
    
    For k = 0 To UBound(arr, 2) - 1
        arr(i, k) = 0   '清空当前保存了的切割方案
    Next k
    'j长度的切割1次
    arr(i, j - 1) = 1
    
    '加上右边的切割方案
    For k = 0 To UBound(arr, 2) - 1
        arr(i, k) = arr(i, k) + arr(i - j, k)
    Next k
End Function
