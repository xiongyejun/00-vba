Attribute VB_Name = "ģ��1"
'��������ԯ�����д�����Excel���ʼ����ƫ���������Ȥ���ɼ���QQȺ141616426����Ⱥ��������֤���룺Excel���ʼ�
Sub CDOSENDEMAIL()
    Dim CDOMail As Variant
    
    On Error Resume Next                                         '��������ִ��
    Application.DisplayAlerts = False                            '����ϵͳ��ʾ
    ThisWorkbook.ChangeFileAccess Mode:=xlReadOnly               '������������Ϊֻ��ģʽ
    
    Set CDOMail = CreateObject("CDO.Message")                    '��������
    CDOMail.From = "123@qq.com"                              '���÷����˵�����
    CDOMail.To = "321@qq.com"                                '���������˵�����
    CDOMail.Subject = "����:��CDO�����ʼ�����"                   '�趨�ʼ�������
    'CDOMail.TextBody = "�ı�����"                               'ʹ���ı���ʽ�����ʼ�
    CDOMail.HtmlBody = "���������˷��ʼ�������CDO������ȷ"       'ʹ��Html��ʽ�����ʼ�
    CDOMail.AddAttachment ThisWorkbook.FullName                  '���ͱ�������Ϊ����
    STUl = "http://schemas.microsoft.com/cdo/configuration/"     '΢���������ַ
    
    With CDOMail.Configuration.Fields
        .Item(STUl & "smtpserver") = "smtp.qq.com"               'SMTP��������ַ
        .Item(STUl & "smtpserverport") = 465                      'SMTP�������˿�
        .Item(STUl & "sendusing") = 2                            '���Ͷ˿�
        .Item(STUl & "smtpauthenticate") = 1                     'Զ�̷�������Ҫ��֤
        .Item(STUl & "sendusername") = "123"                 '���ͷ���������
        .Item(STUl & "sendpassword") = "111"                '���ͷ���������
        .Item(STUl & "smtpconnectiontimeout") = 60               '���ӳ�ʱ���룩
        .Update
    End With
    
    CDOMail.Send                                                  'ִ�з���
    Set CDOMail = Nothing                                         '���ͳɹ���ʱ�ͷŶ���
    If Err.Number = 0 Then
        MsgBox "�ɹ������ʼ�", , "��ܰ��ʾ"                           '���û�г�������ʾ���ͳɹ�
    Else
        MsgBox Err.Description, vbInformation, "�ʼ�����ʧ��"         '�����������ʾ�������ͺʹ������
    End If
    
    ThisWorkbook.ChangeFileAccess Mode:=xlReadWrite               '������������Ϊ��дģʽ
    Application.DisplayAlerts = True                              '�ָ�ϵͳ��ʾ
End Sub

