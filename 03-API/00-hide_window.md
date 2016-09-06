[http://club.excelhome.net/forum.php?mod=viewthread&tid=1261188&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=1261188&extra=page%3D1)

# 当设置EXCEL为隐藏时，如何使用户窗体在Windows工具栏里有图标？ #

	Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
	Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
	Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
	Private Const GWL_EXSTYLE = (-20)
	Private Const WS_EX_APPWINDOW = &H40000
	
	Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
	Private Const SW_SHOW = 5
	Private Const SW_HIDE = 0
	
	Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
	Private Const WM_SETICON = &H80
	Private Const ICON_SMALL = 0
	Private Const ICON_BIG = 1
	
	Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
	
	
	Private Sub UserForm_Activate()
	    Dim lHwnd As Long, lExStyle As Long
	    lHwnd = FindWindow(vbNullString, Me.Caption)
	    lExStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)
	    lExStyle = lExStyle Or WS_EX_APPWINDOW
	    ShowWindow lHwnd, SW_HIDE
	    SetWindowLong lHwnd, GWL_EXSTYLE, lExStyle
	    ShowWindow lHwnd, SW_SHOW
	    
	    SendMessage lHwnd, WM_SETICON, ICON_BIG, ByVal Image1.Picture.handle
	    SendMessage lHwnd, WM_SETICON, ICON_SMALL, ByVal Image1.Picture.handle
	    DrawMenuBar lHwnd
	    Application.Visible = False
	End Sub
	
	
	Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
	    Application.Visible = True
	End Sub