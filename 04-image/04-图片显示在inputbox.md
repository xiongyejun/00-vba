[http://club.excelhome.net/forum.php?mod=viewthread&tid=917414&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=917414&extra=page%3D1)

	Option Explicit
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>功  能：显示获取验证码<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>函数名：mInputBox<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>参  数：PictureData，传入图片字节<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>作  者：hyy514<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>日  期：2012.09.08<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'Hook\subClass--------------------------------------------------------------------------------------------------------
	Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
	Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
	Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
	Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
	Private Const WH_CBT = 5
	Private Const HCBT_ACTIVATE = 5
	
	Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
	Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
	Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
	Private Const GWL_STYLE = (-16)
	Private Const WS_SYSMENU = &H80000
	Private Const GWL_WNDPROC = (-4)
	Private Const WM_CTLCOLORBTN = &H135
	Private Const WM_CTLCOLOREDIT = &H133
	
	Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
	Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
	Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
	Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
	
	'绘图-------------------------------------------------------------------------------------------------------------------
	Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
	Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As GpStatus
	Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Single, ByVal y As Single) As GpStatus
	Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
	Private Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hwnd As Long, graphics As Long) As GpStatus
	Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDc As Long, graphics As Long) As GpStatus
	Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
	Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, image As Long) As GpStatus
	Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As GpStatus
	Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As GpStatus
	Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As GpStatus
	
	Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Long, ByRef image As Long) As Long
	Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
	Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
	
	Private Type GdiplusStartupInput
	    GdiplusVersion As Long
	    DebugEventCallback As Long
	    SuppressBackgroundThread As Long
	    SuppressExternalCodecs As Long
	End Type
	Private Enum GpStatus
	    ok = 0
	    GenericError = 1
	    InvalidParameter = 2
	    OutOfMemory = 3
	    ObjectBusy = 4
	    InsufficientBuffer = 5
	    NotImplemented = 6
	    Win32Error = 7
	    WrongState = 8
	    Aborted = 9
	    FileNotFound = 10
	    ValueOverflow = 11
	    AccessDenied = 12
	    UnknownImageFormat = 13
	    FontFamilyNotFound = 14
	    FontStyleNotFound = 15
	    NotTrueTypeFont = 16
	    UnsupportedGdiplusVersion = 17
	    GdiplusNotInitialized = 18
	    PropertyNotFound = 19
	    PropertyNotSupported = 20
	End Enum
	
	Dim gdip_Token      As Long
	Dim gdip_pngImage   As Long
	Dim gdip_Graphics   As Long
	'--------------------------
	Dim sCaption        As String   '标题
	Dim hHook           As Long     '钩子地址
	Dim mPictureData()  As Byte     '图片数据
	Dim mPicX           As Long     'X座标
	Dim mPicY           As Long     'Y座标
	Dim hwndInput       As Long     'Inputbox主窗口句柄
	Dim oldProcInput    As Long     'Inputbox主窗子类化地址
	Dim bMsgAnd         As Boolean  '并发消息
	
	Public Function mInputBox(ByRef PictureData() As Byte, Optional ByVal picX As Long = 10, Optional ByVal picY As Long = 10)
	    Dim hInst As Long
	    Dim hreadId As Long
	    mPictureData = PictureData
	    mPicX = picX
	    mPicY = picY
	    hInst = Application.Hinstance
	    hreadId = GetCurrentThreadId
	    hHook = SetWindowsHookEx(WH_CBT, AddressOf HookProc, ByVal hInst, ByVal hreadId)
	   
	    sCaption = "请输入验证码:"
	    mInputBox = InputBox("", sCaption)
	    If hHook <> 0 Then
	        UnhookWindowsHookEx hHook
	        hHook = 0
	    End If
	    SetWindowLong hwndInput, GWL_WNDPROC, oldProcInput
	    oldProcInput = 0
	    hwndInput = 0
	    Erase mPictureData
	End Function
	
	Private Function HookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
	    Dim lStyle As Long
	    Dim hwndOK As Long
	    Dim hwndCancel As Long
	    If nCode = HCBT_ACTIVATE Then
	         If GetWinText(wParam) = sCaption And GetWinClassName(wParam) = "#32770" Then
	             lStyle = GetWindowLong(wParam, GWL_STYLE)
	             lStyle = lStyle And Not WS_SYSMENU
	             SetWindowLong wParam, GWL_STYLE, lStyle
	           
	             hwndOK = FindWindowEx(wParam, 0, "Button", vbNullString)
	             hwndCancel = FindWindowEx(wParam, hwndOK, "Button", vbNullString)
	             ShowWindow hwndCancel, False
	             hwndInput = wParam
	             oldProcInput = SetWindowLong(hwndInput, GWL_WNDPROC, AddressOf InputProc)
	             HookProc = CallNextHookEx(hHook, nCode, wParam, lParam)
	             UnhookWindowsHookEx hHook
	             hHook = 0
	             Exit Function
	         End If
	    End If
	    HookProc = CallNextHookEx(hHook, nCode, wParam, lParam)
	    
	End Function
	
	Private Function InputProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
	    Select Case Msg
	        Case WM_CTLCOLOREDIT
	            bMsgAnd = True
	        Case WM_CTLCOLORBTN
	            If bMsgAnd Then
	               bMsgAnd = False
	               PainPic mPictureData, hwnd, mPicX, mPicY
	            End If
	    End Select
	    InputProc = CallWindowProc(oldProcInput, hwnd, Msg, wParam, lParam)
	End Function
	
	Private Function GetWinClassName(hwnd As Long) As String
	    Dim sBuffer As String
	    sBuffer = Space(255)
	    GetClassName hwnd, sBuffer, 255
	    GetWinClassName = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
	End Function
	
	Private Function GetWinText(hwnd As Long) As String
	    Dim sBuffer As String
	    sBuffer = Space(255)
	    GetWindowText hwnd, ByVal sBuffer, 255
	    GetWinText = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
	End Function
	
	Private Sub PainPic(ByRef PictureData() As Byte, ByVal hwnd As Long, ByVal mX As Long, ByVal mY As Long)
	    Dim lngHeight       As Long
	    Dim lngWidth        As Long
	    Dim StreamObject    As Long
	    Dim hDc             As Long
	    
	    Call GDI_Initialize
	    hDc = GetDC(hwnd)
	    If GdipCreateFromHDC(hDc, gdip_Graphics) <> ok Then
	        GdiplusShutdown gdip_Token
	    Else
	        Call CreateStreamOnHGlobal(PictureData(0), 0, StreamObject)
	        Call GdipLoadImageFromStream(StreamObject, gdip_pngImage)
	        Call GdipGetImageHeight(gdip_pngImage, lngHeight)   '
	        Call GdipGetImageWidth(gdip_pngImage, lngWidth)
	        Call GdipDrawImageRect(gdip_Graphics, gdip_pngImage, mX, mY, lngWidth, lngHeight)
	        'Debug.Print gdip_Graphics, gdip_pngImage, mX, mY, lngWidth, lngHeight
	    End If
	    Call GDI_Terminate
	End Sub
	
	Private Sub GDI_Initialize()
	    Dim GpInput As GdiplusStartupInput
	    Dim ret As GpStatus
	    GpInput.GdiplusVersion = 1
	    gdip_Token = 0
	    gdip_pngImage = 0
	    gdip_Graphics = 0
	    ret = GdiplusStartup(gdip_Token, GpInput)
	    If ret <> ok Then
	        Debug.Print "GDI初始化失败！"
	    End If
	End Sub
	
	Private Sub GDI_Terminate()
	    GdipDisposeImage gdip_pngImage
	    GdipDeleteGraphics gdip_Graphics
	    GdiplusShutdown gdip_Token
	    gdip_Token = 0
	    gdip_pngImage = 0
	    gdip_Graphics = 0
	End Sub


	Sub try()
	    Dim Xml As Object
	    Dim picAry() As Byte
	    Set Xml = CreateObject("Microsoft.XMLHTTP")
	    Xml.Open "GET", "https://ssl.captcha.qq.com/getimage", False
	    Xml.Send
	    picAry = Xml.responseBody
	    MsgBox mInputBox(picAry) '腾讯的验证码
	    
	    Xml.Open "GET", "https://webim.feixin.10086.cn/WebIM/GetPicCode.aspx?Type=ccpsession" & Rnd, False
	    Xml.Send
	    picAry = Xml.responseBody
	    MsgBox mInputBox(picAry) '飞信的
	    
	End Sub
