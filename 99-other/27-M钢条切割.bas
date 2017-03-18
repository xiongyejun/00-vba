Attribute VB_Name = "M�����и�"
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

'ֱ�Ӱ��� �۸�/�� �������У�ֱ�����

'�㷨���ԣ�������
'����            3-8Ԫ��2-3Ԫ��8-20Ԫ
'�и��        8
'����            3-8Ԫ-2.67Ԫ/�ס�8-20Ԫ-2.5Ԫ/�ס�2-3Ԫ-1.5Ԫ/��
'���            3�׵�2����2�׵�1�������19Ԫ������8-20Ԫ
'ԭ��            δ��ֿ��Ǳ�ƽ�������
Function vba_main_2()
    Dim arr_price() As Long '0��-���� 1��-�۸�
    Dim arr_sort() As Double
    Dim i_col As Long
    Dim arr_result() As Long
    Dim i As Long
    Dim tmp_len As Long
    Dim p_price As Long
    
    '��ȡ��֪��ÿ�����ȶ�Ӧ�ļ۸�
    i_col = Range("A1").End(xlToRight).Column
    ReDim arr_price(i_col - 1, 1) As Long
    ReDim arr_sort(i_col - 1) As Double
    
    For i = 2 To i_col
        arr_price(i - 1, 0) = Cells(1, i).Value
        arr_price(i - 1, 1) = Cells(2, i).Value
        arr_sort(i - 1) = VBA.CDbl(Cells(2, i).Value / Cells(1, i).Value)
    Next i
    '���յ�λ�۸�������
    InsertSort arr_sort, 0, i_col - 1, arr_price
    
    ReDim arr_result(1 To 8, 1 To i_col) As Long
    For i = 1 To 8
        tmp_len = Cells(i + 13, 1).Value
        p_price = 0
        
        Do Until tmp_len = 0
            
            '������ǰ��ĳ������и�
            Do While arr_price(p_price, 0) <= tmp_len
                arr_result(i, i_col) = arr_result(i, i_col) + arr_price(p_price, 1) '���
                arr_result(i, arr_price(p_price, 0)) = arr_result(i, arr_price(p_price, 0)) + 1 '��ǰ�����и�����
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
            ShaoBing = arr_sort(i)             '�����ڱ�
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
    Dim arr_num() As Long   'ÿ�����ȵ��и�����
    Dim arr() As Long
    Dim N As Long
    Dim best_len As Long    '��ֵ����õģ�Ҳ���ǲ��ʱ��������
    Dim d_best As Double
    
    '��ȡ��֪��ÿ�����ȶ�Ӧ�ļ۸�
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
    
    'n��0�У���Ӧÿ�����ȵ����Ž�۸�
    '������ж�Ӧ���ŵ��и�����
    'ֻҪ���һ�ξͿ��ԣ������Ķ����������
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
'arr_price    �±곤�ȶ�Ӧ�ļ۸�
'arr          ��ʱ�洢ÿ�����ȵ����Ž�arr(n,0)��ͬʱ�洢ÿ������и������ÿ�ּ۸���Ҫ�и�ĸ���
'n            �����ĳ���
'arr_num      ÿ���۸񳤶���Ҫ�и�ĸ���
'best_len     '��ֵ����õģ�Ҳ���ǲ��ʱ��������
Function BottomUpCutRod(arr_price() As Long, N As Long, arr() As Long, arr_num() As Long, best_len As Long) As Long
    Dim i As Long, j As Long
    Dim i_max As Long
    Dim i_tmp As Long
    Dim max_price_col As Long
    
    max_price_col = UBound(arr_price)
    '����������м۸����󳤶ȣ�ֻ�ܲ��
    If N > max_price_col Then
        Do While N > max_price_col
            N = N - best_len
            BottomUpCutRod = BottomUpCutRod(arr_price, best_len, arr, arr_num, best_len) + BottomUpCutRod(arr_price, N, arr, arr_num, best_len)
        Loop
        Exit Function
    End If
    
    '����ǳ����˵ĳ��ȣ����п����Ѿ������arr��ֱ�Ӹ�ֵ�Ϳ���
    If arr(N, 0) > 0 Then GoTo lable_quit
    
    '������������и��³���Ϊi��һ�Σ�ֻ���ұ�ʣ�µĳ���Ϊn-i��һ�μ��������и����ߵ�һ�����ٽ����и�
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
    '����Ҫ����������Ϊ�п����ǲ�ֹ�����
    For i = 1 To UBound(arr_num)
        arr_num(i) = arr_num(i) + arr(N, i)
    Next i

    BottomUpCutRod = arr(N, 0)
End Function
'��¼ÿ���۸񳤶ȵ����Ž��и��
'i  ��ǰ�и�ĸ�������
'j  ����и�Ϊj�ĳ���
Function GetCutNum(arr() As Long, i As Long, j As Long)
    Dim k As Long
    
    For k = 1 To UBound(arr, 2)
        arr(i, k) = 0   '��յ�ǰ�����˵��и��
    Next k
    'j���ȵ��и�1��
    arr(i, j) = 1
    
    '�����ұߵ��и��
    For k = 1 To UBound(arr, 2)
        arr(i, k) = arr(i, k) + arr(i - j, k)
    Next k
End Function
