Attribute VB_Name = "MCombin"
Option Explicit

'��1��DataNum����һά�����У���ѡChooseNum������Ͻ��

Type CombinResultItem
    Arr() As String     '1���������1��chooseNum����������
End Type

Type CombinType
    Data() As String    'һ������Դ
    DataNum As Long     '����Դ�ĸ���
    ChooseNum As Long   'Ҫѡ����ٸ������
    
    Result() As CombinResultItem  '���
    ResultNum As Long   '����ĸ����������ù�������Combin����
    pResult As Long     'ָ���������ɵĽ��
End Type

Sub test()
    Dim cbType As CombinType
    Dim i As Long
    Const NUM_DATA As Long = 10
    
    '��ʼ������
    ReDim cbType.Data(NUM_DATA - 1) As String
    For i = 0 To NUM_DATA - 1
        cbType.Data(i) = VBA.CStr(i)
    Next i
    
    cbType.DataNum = NUM_DATA
    cbType.ChooseNum = 2
    cbType.ResultNum = Application.WorksheetFunction.Combin(NUM_DATA, cbType.ChooseNum)
    
    '��ʼ���������
    ReDim cbType.Result(cbType.ResultNum - 1) As CombinResultItem
    For i = 0 To cbType.ResultNum - 1
        ReDim cbType.Result(i).Arr(cbType.ChooseNum - 1) As String
    Next
    '��ʼ���
    DGCombin cbType, 0, 0
    '��ӡ��Ͻ��
    For i = 0 To cbType.ResultNum - 1
        Debug.Print i, VBA.Join(cbType.Result(i).Arr, "��")
    Next
End Sub
'pData        'ָ����Ҫʹ�õ�����Դ�±�
'pChooseNum   '��ϵ��ڼ�����
Function DGCombin(cbType As CombinType, pData As Long, pChooseNum As Long)
    Dim i As Long
    
    If pChooseNum = cbType.ChooseNum Then
        cbType.pResult = cbType.pResult + 1
        Exit Function
    End If
    
    cbType.Result(cbType.pResult).Arr(pChooseNum) = cbType.Data(pData)
    DGCombin cbType, pData + 1, pChooseNum + 1
    
    'ʣ�����ݵĸ���������ʣ�»���Ҫ�����ݸ������������������
    If cbType.DataNum - pData > cbType.ChooseNum - pChooseNum Then
        'pChooseNum ֮ǰ��������Ҫ���ƹ���
        For i = 0 To pChooseNum - 1
            cbType.Result(cbType.pResult).Arr(i) = cbType.Result(cbType.pResult - 1).Arr(i)
        Next
        DGCombin cbType, pData + 1, pChooseNum
    End If
End Function

