Attribute VB_Name = "M�����и�"
Option Explicit

Sub vba_main()
    Dim i_col As Long
    Dim arr_price() As Long
    Dim i As Long
    Dim arr_num() As Long   'ÿ�����ȵ��и�����
    Dim arr() As Long
    Dim n As Long
    
    '��ȡ��֪��ÿ�����ȶ�Ӧ�ļ۸�
    i_col = Range("A1").End(xlToRight).Column
    ReDim arr_price(i_col - 1) As Long
    For i = 2 To i_col
        arr_price(i - 1) = Cells(2, i).Value
    Next i
    
    'n�����1�У���Ӧÿ�����ȵ����Ž�۸�
    'ǰ����ж�Ӧ���ŵ��и�����
    'ֻҪ���һ�ξͿ��ԣ������Ķ����������
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
'arr_price    �±곤�ȶ�Ӧ�ļ۸�
'arr          ��ʱ�洢ÿ�����ȵ����Ž�arr(n,0)��ͬʱ�洢ÿ������и������ÿ�ּ۸���Ҫ�и�ĸ���
'n            �����ĳ���
'arr_num      ÿ���۸񳤶���Ҫ�и�ĸ���
Function BottomUpCutRod(arr_price() As Long, n As Long, arr() As Long) As Long
    Dim i As Long, j As Long
    Dim i_max As Long
    Dim i_tmp As Long
    Dim max_price_col As Long
    
    max_price_col = UBound(arr_price)
    
    If arr(n, max_price_col) > 0 Then GoTo lable_quit
    
    '������������и��³���Ϊi��һ�Σ�ֻ���ұ�ʣ�µĳ���Ϊn-i��һ�μ��������и����ߵ�һ�����ٽ����и�
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
'��¼ÿ���۸񳤶ȵ����Ž��и��
'i  ��ǰ�и�ĸ����ĳ���
'j  ����и�Ϊj�ĳ���
Function GetCutNum(arr() As Long, i As Long, j As Long)
    Dim k As Long
    
    For k = 0 To UBound(arr, 2) - 1
        arr(i, k) = 0   '��յ�ǰ�����˵��и��
    Next k
    'j���ȵ��и�1��
    arr(i, j - 1) = 1
    
    '�����ұߵ��и��
    For k = 0 To UBound(arr, 2) - 1
        arr(i, k) = arr(i, k) + arr(i - j, k)
    Next k
End Function
