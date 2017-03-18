Attribute VB_Name = "M�����и�"
Option Explicit

Sub vba_main()
    Dim i_col As Long
    Dim arr_price() As Long
    Dim i As Long
    Dim arr_num() As Long   'ÿ�����ȵ��и�����
    Dim arr() As Long
    Dim n As Long
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
            n = Cells(i, 1).Value
            ReDim arr_num(1 To Cells(1, i_col).Value) As Long
            
            .Offset(0, 10).Value = BottomUpCutRod(arr_price, n, arr, arr_num, best_len)
            .Resize(1, i_col - 1).Value = arr_num
        End With
    Next i
End Sub
'arr_price    �±곤�ȶ�Ӧ�ļ۸�
'arr          ��ʱ�洢ÿ�����ȵ����Ž�arr(n,0)��ͬʱ�洢ÿ������и������ÿ�ּ۸���Ҫ�и�ĸ���
'n            �����ĳ���
'arr_num      ÿ���۸񳤶���Ҫ�и�ĸ���
'best_len     '��ֵ����õģ�Ҳ���ǲ��ʱ��������
Function BottomUpCutRod(arr_price() As Long, n As Long, arr() As Long, arr_num() As Long, best_len As Long) As Long
    Dim i As Long, j As Long
    Dim i_max As Long
    Dim i_tmp As Long
    Dim max_price_col As Long
    
    max_price_col = UBound(arr_price)
    '����������м۸����󳤶ȣ�ֻ�ܲ��
    If n > max_price_col Then
        Do While n > max_price_col
            n = n - best_len
            BottomUpCutRod = BottomUpCutRod(arr_price, best_len, arr, arr_num, best_len) + BottomUpCutRod(arr_price, n, arr, arr_num, best_len)
        Loop
        Exit Function
    End If
    
    '����ǳ����˵ĳ��ȣ����п����Ѿ������arr��ֱ�Ӹ�ֵ�Ϳ���
    If arr(n, 0) > 0 Then GoTo lable_quit
    
    '������������и��³���Ϊi��һ�Σ�ֻ���ұ�ʣ�µĳ���Ϊn-i��һ�μ��������и����ߵ�һ�����ٽ����и�
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
    '����Ҫ����������Ϊ�п����ǲ�ֹ�����
    For i = 1 To UBound(arr_num)
        arr_num(i) = arr_num(i) + arr(n, i)
    Next i

    BottomUpCutRod = arr(n, 0)
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
