Attribute VB_Name = "M钢条切割"
Option Explicit

Sub vba_main()
    Dim i_col As Long
    Dim arr_price() As Long
    Dim i As Long
    Dim arr_num() As Long   '每个长度的切割数量
    Dim arr() As Long
    Dim n As Long
    Dim best_len As Long    '价值量最好的，也就是拆分时按这个拆分
    Dim d_best As Double
    
    '获取已知的每个长度对应的价格
    i_col = Range("A1").End(xlToRight).Column
    ReDim arr_price(i_col - 1) As Long
    d_best = 0#
    For i = 2 To i_col
        arr_price(i - 1) = Cells(2, i).Value
        If VBA.CDbl(Cells(2, i).Value / Cells(1, i).Value) > d_best Then
            d_best = VBA.CDbl(Cells(2, i).Value / Cells(1, i).Value)
            best_len = Cells(1, i).Value
        End If
    Next i
    
    'n行0列，对应每个长度的最优解价格
    '后面的列对应最优的切割数量
    '只要求解一次就可以，其他的都可以用这个
    ReDim arr(Cells(1, i_col).Value, i_col - 1) As Long
    
    For i = 4 To 11
        With Cells(i, 2)
            n = Cells(i, 1).Value
            ReDim arr_num(1 To Cells(1, i_col).Value) As Long
            
            .Offset(0, 10).Value = BottomUpCutRod(arr_price, n, arr, arr_num, best_len)
            .Resize(1, i_col - 1).Value = arr_num
        End With
    Next i
End Sub
'arr_price    下标长度对应的价格
'arr          临时存储每个长度的最优解arr(n,0)，同时存储每个解的切割方案：即每种价格需要切割的个数
'n            钢条的长度
'arr_num      每个价格长度需要切割的个数
'best_len     '价值量最好的，也就是拆分时按这个拆分
Function BottomUpCutRod(arr_price() As Long, n As Long, arr() As Long, arr_num() As Long, best_len As Long) As Long
    Dim i As Long, j As Long
    Dim i_max As Long
    Dim i_tmp As Long
    Dim max_price_col As Long
    
    max_price_col = UBound(arr_price)
    '如果超过了有价格的最大长度，只能拆分
    If n > max_price_col Then
        Do While n > max_price_col
            n = n - best_len
            BottomUpCutRod = BottomUpCutRod(arr_price, best_len, arr, arr_num, best_len) + BottomUpCutRod(arr_price, n, arr, arr_num, best_len)
        Loop
        Exit Function
    End If
    
    '如果是超过了的长度，很有可能已经求出了arr，直接赋值就可以
    If arr(n, 0) > 0 Then GoTo lable_quit
    
    '将钢条从左边切割下长度为i的一段，只对右边剩下的长度为n-i的一段继续进行切割，对左边的一段则不再进行切割
    For i = 1 To n
        i_max = -1
        
        For j = 1 To i
            i_tmp = arr_price(j) + arr(i - j, 0)
            
            If i_max < i_tmp Then
                i_max = i_tmp
                GetCutNum arr, i, j
            End If
            
        Next j
        arr(i, 0) = i_max
    Next i
    
lable_quit:
    '这里要加起来，因为有可能是拆分过来的
    For i = 1 To UBound(arr_num)
        arr_num(i) = arr_num(i) + arr(n, i)
    Next i

    BottomUpCutRod = arr(n, 0)
End Function
'记录每个价格长度的最优解切割方案
'i  当前切割的钢条长度
'j  左边切割为j的长度
Function GetCutNum(arr() As Long, i As Long, j As Long)
    Dim k As Long
    
    For k = 1 To UBound(arr, 2)
        arr(i, k) = 0   '清空当前保存了的切割方案
    Next k
    'j长度的切割1次
    arr(i, j) = 1
    
    '加上右边的切割方案
    For k = 1 To UBound(arr, 2)
        arr(i, k) = arr(i, k) + arr(i - j, k)
    Next k
End Function
