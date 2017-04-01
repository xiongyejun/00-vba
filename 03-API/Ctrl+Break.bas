Attribute VB_Name = "模块1"
Option Explicit

Declare Sub keybd_event Lib "user32" ( _
                                ByVal bVk As Byte, _
                                ByVal bScan As Byte, _
                                ByVal dwFlags As Long, _
                                ByVal dwExtraInfo As Long)

Const KEYEVENT_KEYUP = &H2
Const VK_CANCEL = &H3 'Control-break processing

Sub ctrl_break()
    
    Call keybd_event(VK_CANCEL, 0, 0, 0)       '按下
    Call keybd_event(VK_CANCEL, 0, KEYEVENT_KEYUP, 0)    '松开
    
End Sub

Sub test()
    Const 窗口标题 = "模板.xlsm" ' 这个标题改为具体的Excel窗口标题！
    AppActivate 窗口标题
    ctrl_break
End Sub

