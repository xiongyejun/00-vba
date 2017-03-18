Attribute VB_Name = "M钢条切割"
Option Explicit

Sub vba_main()
    Dim i As Long
    Dim t As Double
    Const N As Long = 1
    
    t = Timer
    For i = 1 To N
        vba_main_1
    Next i
    Range("O3").Value = Timer - t
    
    t = Timer
    For i = 1 To N
        vba_main_2
    Next i
    Range("O13").Value = Timer - t
End Sub

'直接按照 价格/米 降序排列，直接输出

'算法不对，有问题
'比如            3-8元、2-3元、8-20元
'切割长度        8
'排序            3-8元-2.67元/米、8-20元-2.5元/米、2-3元-1.5元/米
'结果            3米的2个、2米的1个，金额19元，低于8-20元
'原因            未充分考虑被平均的情况
Function vba_main_2()
    Dim arr_price() As Long '0列-长度 1列-价格
    Dim arr_sort() As Double
    Dim i_col As Long
    Dim arr_result() As Long
    Dim i As Long
    Dim tmp_len As Long
    Dim p_price As Long
    
    '获取已知的每个长度对应的价格
    i_col = Range("A1").End(xlToRight).Column
    ReDim arr_price(i_col - 1, 1) As Long
    ReDim arr_sort(i_col - 1) As Double
    
    For i = 2 To i_col
        arr_price(i - 1, 0) = Cells(1, i).Value
        arr_price(i - 1, 1) = Cells(2, i).Value
        arr_sort(i - 1) = VBA.CDbl(Cells(2, i).Value / Cells(1, i).Value)
    Next i
    '按照单位价格降序排列
    InsertSort arr_sort, 0, i_col - 1, arr_price
    
    ReDim arr_result(1 To 8, 1 To i_col) As Long
    For i = 1 To 8
        tmp_len = Cells(i + 13, 1).Value
        p_price = 0
        
        Do Until tmp_len = 0
            
            '按照最前面的长度来切割
            Do While arr_price(p_price, 0) <= tmp_len
                arr_result(i, i_col) = arr_result(i, i_col) + arr_price(p_price, 1) '金额
                arr_result(i, arr_price(p_price, 0)) = arr_result(i, arr_price(p_price, 0)) + 1 '当前长度切割数量
                tmp_len = tmp_len - arr_price(p_price, 0)
            Loop
            p_price = p_price + 1
        Loop
    Next i
    
    Range("B14:L21").Value = arr_result
End Function

Function InsertSort(arr_sort() As Double, Low As Long, High As Long, arr_price() As Long)
    Dim i As Long, j As Long
    Dim ShaoBing As Double
    Dim ShaoBing_0 As Double
    Dim ShaoBing_1 As Double
    
    For i = Low + 1 To High
    
        If arr_sort(i) > arr_sort(i - 1) Then
            ShaoBing = arr_sort(i)             '设置哨兵
            ShaoBing_0 = arr_price(i, 0)
            ShaoBing_1 = arr_price(i, 1)
            
            j = i - 1
            Do While arr_sort(j) < ShaoBing
                arr_sort(j + 1) = arr_sort(j)
                arr_price(j + 1, 0) = arr_price(j, 0)
                arr_price(j + 1, 1) = arr_price(j, 1)
                
                j = j - 1
                If j = Low - 1 Then Exit Do
            Loop
            
            arr_sort(j + 1) = ShaoBing
            arr_price(j + 1, 0) = ShaoBing_0
            arr_price(j + 1, 1) = ShaoBing_1
        End If
    
    Next i
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function vba_main_1()
    Dim i_col As Long
    Dim arr_price() As Long
    Dim i As Long
    Dim arr_num() As Long   '每个长度的切割数量
    Dim arr() As Long
    Dim N As Long
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
            N = Cells(i, 1).Value
            ReDim arr_num(1 To Cells(1, i_col).Value) As Long
            
            .Offset(0, 10).Value = BottomUpCutRod(arr_price, N, arr, arr_num, best_len)
            .Resize(1, i_col - 1).Value = arr_num
        End With
    Next i
End Function
'arr_price    下标长度对应的价格
'arr          临时存储每个长度的最优解arr(n,0)，同时存储每个解的切割方案：即每种价格需要切割的个数
'n            钢条的长度
'arr_num      每个价格长度需要切割的个数
'best_len     '价值量最好的，也就是拆分时按这个拆分
Function BottomUpCutRod(arr_price() As Long, N As Long, arr() As Long, arr_num() As Long, best_len As Long) As Long
    Dim i As Long, j As Long
    Dim i_max As Long
    Dim i_tmp As Long
    Dim max_price_col As Long
    
    max_price_col = UBound(arr_price)
    '如果超过了有价格的最大长度，只能拆分
    If N > max_price_col Then
        Do While N > max_price_col
            N = N - best_len
            BottomUpCutRod = BottomUpCutRod(arr_price, best_len, arr, arr_num, best_len) + BottomUpCutRod(arr_price, N, arr, arr_num, best_len)
        Loop
        Exit Function
    End If
    
    '如果是超过了的长度，很有可能已经求出了arr，直接赋值就可以
    If arr(N, 0) > 0 Then GoTo lable_quit
    
    '将钢条从左边切割下长度为i的一段，只对右边剩下的长度为n-i的一段继续进行切割，对左边的一段则不再进行切割
    For i = 1 To N
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
        arr_num(i) = arr_num(i) + arr(N, i)
    Next i

    BottomUpCutRod = arr(N, 0)
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
