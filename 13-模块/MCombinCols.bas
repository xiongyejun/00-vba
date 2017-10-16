Attribute VB_Name = "MCombinCols"
Option Explicit

'����Դ�Ƕ��1ά���飬���1ά����֮�����ϣ��ж����о�ȡ���ٸ��������
'������ϵ�����ͬ1����������ݲ��ܳ���2��
Type MyType
    i_rows As Long
    arr_data() As Variant
End Type

Type CombinColsDataType
    Arr() As String     'һ������Դ
    Count As Long       '���ݵĸ���
End Type

Type ReptCountType
    Self As Long        '�����ظ��Ĵ���=���������
    Col As Long         '������Ҫѭ���Ĵ���=��ǰ�������
End Type

Type CombinColsType
    Data() As CombinColsDataType     '��������Դ
    ColNum As Long                   '����Դ���ж�����
    
    Result() As CombinColsDataType   '���
    ReptCount() As ReptCountType     'ѭ����������
    ResultNum As Long                '����ĸ���--���ڸ��и������
    pResult As Long                  'ָ���������ɵĽ��
End Type

Sub testCombinCol()
    Dim cbType As CombinColsType
    Dim i As Long, j As Long
    
    cbType.ColNum = 3
    ReDim cbType.Data(cbType.ColNum - 1) As CombinColsDataType
    '��ʼ������
    cbType.Data(0).Count = 7
    ReDim cbType.Data(0).Arr(cbType.Data(0).Count - 1) As String
    For i = 0 To cbType.Data(0).Count - 1
        cbType.Data(0).Arr(i) = VBA.CStr(i)
    Next
    
    cbType.Data(1).Count = 2
    ReDim cbType.Data(1).Arr(cbType.Data(1).Count - 1) As String
    For i = 0 To cbType.Data(1).Count - 1
        cbType.Data(1).Arr(i) = VBA.Chr(i + VBA.Asc("a"))
    Next
    
    cbType.Data(2).Count = 3
    ReDim cbType.Data(2).Arr(cbType.Data(2).Count - 1) As String
    For i = 0 To cbType.Data(2).Count - 1
        cbType.Data(2).Arr(i) = VBA.Chr(i + VBA.Asc("A"))
    Next
    '����ÿ������Ӧ���ظ��Ĵ���
    ReDim cbType.ReptCount(cbType.ColNum - 1) As ReptCountType
    cbType.ResultNum = 1
    For i = 0 To cbType.ColNum - 1
        cbType.ReptCount(i).Self = 1
        cbType.ReptCount(i).Col = 1
        
        cbType.ResultNum = cbType.ResultNum * cbType.Data(i).Count
        
        '�����ظ��Ĵ���=���������
        For j = i + 1 To cbType.ColNum - 1
            cbType.ReptCount(i).Self = cbType.ReptCount(i).Self * cbType.Data(j).Count
        Next
        
        '������Ҫѭ���Ĵ���=��ǰ�������
        For j = 0 To i - 1
            cbType.ReptCount(i).Col = cbType.ReptCount(i).Col * cbType.Data(j).Count
        Next
    Next
    '��ʼ���������
    ReDim cbType.Result(cbType.ResultNum - 1) As CombinColsDataType
    For i = 0 To cbType.ResultNum - 1
        ReDim cbType.Result(i).Arr(cbType.ColNum - 1) As String
    Next
    
    CombinCols cbType
    For i = 0 To cbType.ResultNum - 1
        Debug.Print i, VBA.Join(cbType.Result(i).Arr, "��")
    Next
End Sub

Function CombinCols(cbType As CombinColsType)
    Dim i As Long, j As Long, k As Long, m As Long

    '���ݽ����index��λ�����ݵ��У�ֱ�Ӹ�ֵ
    For i = 0 To cbType.ResultNum - 1
        For j = 0 To cbType.ColNum - 1
            k = i \ cbType.ReptCount(j).Self
            m = k Mod cbType.Data(j).Count
            
            cbType.Result(i).Arr(j) = cbType.Data(j).Arr(m)
        Next j
    Next i
End Function

