Attribute VB_Name = "ģ��2"
Option Explicit

Sub RowCnt()
    '���С����ֻ����һ�仰��txt�ļ���ȡ���������У�̫�����ˣ���һ���Ĵ��밡��
    t = Timer
    Dim Arr, k&
    fnum = FreeFile '����һ���ڲ���ֵ���
    Open "D:\��������\����.txt" For Input As #fnum
    '�������Ľ��ͣ���Ϊ̫�����ˣ����ò�˵��һ�������Ժ����ˣ�
    '����LOF����ʾ�� Open ���򿪵��ļ��Ĵ�С���ô�С���ֽ�Ϊ��λ��
    'Ȼ��InputB�ٶ��������Ƕ����ƣ����䲻��ת������Ϊ����
    'StrConv ������ָ������ת���������� vbUnicode
    'Split��������һ���±���㿪ʼ��һά���飬������ָ����Ŀ�����ַ������ֿ��ļ�ֵ�ǻس� vbCrLf
    Arr = Split(StrConv(InputB(LOF(fnum), 1), vbUnicode), vbCrLf)  '�����س������������ַ���Ϊ������
    k = UBound(Arr) + 1
    Reset
    Close #fnum
    MsgBox "��ʱ��" & Timer - t
    MsgBox "���ı�������Ϊ" & k
    '����һ��6MB��txt�ļ�ֻ���˲���4��
End Sub

