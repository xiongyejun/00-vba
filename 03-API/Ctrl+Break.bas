Attribute VB_Name = "ģ��1"
Option Explicit

Declare Sub keybd_event Lib "user32" ( _
                                ByVal bVk As Byte, _
                                ByVal bScan As Byte, _
                                ByVal dwFlags As Long, _
                                ByVal dwExtraInfo As Long)

Const KEYEVENT_KEYUP = &H2
Const VK_CANCEL = &H3 'Control-break processing

Sub ctrl_break()
    
    Call keybd_event(VK_CANCEL, 0, 0, 0)       '����
    Call keybd_event(VK_CANCEL, 0, KEYEVENT_KEYUP, 0)    '�ɿ�
    
End Sub

Sub test()
    Const ���ڱ��� = "ģ��.xlsm" ' ��������Ϊ�����Excel���ڱ��⣡
    AppActivate ���ڱ���
    ctrl_break
End Sub

