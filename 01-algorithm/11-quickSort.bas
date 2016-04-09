Attribute VB_Name = "QSort"
Option Explicit
Sub Test()
    Dim Arr(), i As Long
    Dim k As Long, t As Double
    
    k = Range("B65536").End(xlUp).Row
    ReDim Arr(1 To k)
    
    For i = 1 To k
        Arr(i) = Cells(i, 2).Value
    Next i
    
    t = Timer
    Call QuickSort(Arr, 1, k)
    Range("E1").Value = Timer - t
    Range("F1").Value = k & "������"
    
    SetShape "Qsort", Range("D1")
    Range("C1:C" & k).ClearContents
    Range("C1:C" & k).Value = Application.WorksheetFunction.Transpose(Arr)
    Erase Arr
End Sub

Sub QuickSort(l(), Low As Long, High As Long)  '��������
    Dim Pivot As Long
    
    If High - Low > 50 Then
        Do While Low < High
            
            Pivot = MyPartition(l, Low, High)
            
            Call QuickSort(l, Low, Pivot - 1)     '�Ե��ӱ�ݹ�����
            Low = Pivot + 1                       'β�ݹ�
        Loop
    
    Else
        Call InsertSort(l, Low, High)
    End If
End Sub

Function MyPartition(l(), ByVal Low As Long, ByVal High As Long) As Long
    Dim PivotKey        '����
    
    PivotKey = MedianOfThree(l, Low, High) '����ȡ��
    
    Do While Low < High
        Do While Low < High And l(High) >= PivotKey
            High = High - 1
        Loop
'        Call Swap(L, Low, High)  '���������¼С�ļ�¼�������Ͷ�
        l(Low) = l(High)        '�����滻�����ǽ����ķ�ʽ���в���
        
        Do While Low < High And l(Low) <= PivotKey
            Low = Low + 1
        Loop
'        Call Swap(L, Low, High)  '���������¼��ļ�¼�������߶�
        l(High) = l(Low)
        
    Loop
    
    l(Low) = PivotKey
    MyPartition = Low
End Function

Function Swap(l(), Low As Long, High As Long)
    Dim iTemp
    iTemp = l(Low)
    l(Low) = l(High)
    l(High) = iTemp
End Function

Private Function MedianOfThree(l(), ByVal Low As Long, ByVal High As Long)
    Dim m As Long
    
    m = Low + (High - Low) / 2
    
    If l(Low) > l(High) Then Call Swap(l, Low, High) '����������Ҷ����ݣ���֤��˽�С
    If l(m) > l(High) Then Call Swap(l, High, m)     '�����м����Ҷ����ݣ���֤�м��С
    If l(m) > l(Low) Then Call Swap(l, m, Low)       '�����м���������ݣ���֤���Ϊ�м�ֵ
    MedianOfThree = l(Low)
    
End Function
Sub SetShape(ButtonStr As String, Rng As Range)
    With ActiveSheet.Shapes(ButtonStr)
        .Left = Rng.Left
        .Width = Rng.Width
        .Height = Rng.Height
        .Top = Rng.Top
    End With
End Sub

Sub TestMedianOfThree()
    Dim Arr(1 To 3)
    Arr(1) = 11
    Arr(2) = 31
    Arr(3) = 21
    MedianOfThree Arr, 1, 3

End Sub
