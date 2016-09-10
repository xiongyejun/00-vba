
[http://club.excelhome.net/thread-1257381-1-1.html](http://club.excelhome.net/thread-1257381-1-1.html)

# 控件直接显示网络图片 #

	'内存函数
	Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
	Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
	Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
	Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
	Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
	
	'OLE函数
	Public Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
	Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
	Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
	Private Type PicBmp
	    Size   As Long
	    Type   As Long
	    hBmp   As Long
	    hPal   As Long
	    Reserved   As Long
	End Type
	
	'GDIplus函数
	Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
	Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
	Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
	Private Declare Function GdipLoadImageFromStream Lib "GDIPlus" (ByVal stream As IUnknown, Image As Long) As Long
	Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long
	
	Private Type GdiplusStartupInput
	    GdiplusVersion As Long
	    DebugEventCallback As Long
	    SuppressBackgroundThread As Long
	    SuppressExternalCodecs As Long
	End Type
	
	Private Const GMEM_MOVEABLE = &H2
	
	Public Function LoadWebImage(url As String) As StdPicture
	    
	    Dim hMem As Long
	    Dim nSize As Long
	    Dim lpData As Long
	    Dim bufferBytes() As Byte
	    Dim istm As stdole.IUnknown
	    Dim lToken As Long
	    Dim lGSI As GdiplusStartupInput
	    Dim IID_IDispatch(15) As Byte
	    Dim pic As PicBmp
	    Dim lImage As Long, hBmp As Long
	    
	    Dim httpRequest
	    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")            '创建WinHttpRequest对象
	    With httpRequest
	        .Open "get", url, False                                             '获取URL内容
	        .Send
	        If Left(.GetResponseHeader("Content-Type"), 6) = "image/" Then      '如果URL为图片
	            '############nSize = .GetResponseHeader("Content-Length")'获取网络图片的字节长度
            	'##########ReDim bufferBytes(nSize - 1)
            	bufferBytes = .ResponseBody'将图片文件存储到字节数组中
           		nSize = UBound(bufferBytes) + 1  '获取网络图片的字节###########
	            bufferBytes = .ResponseBody                                     '将图片文件存储到字节数组中
	            hMem = GlobalAlloc(GMEM_MOVEABLE, nSize)                        '分配一块全局内存
	            lpData = GlobalLock(hMem)                                       '获取内存句柄
	            CopyMemory lpData, VarPtr(bufferBytes(0)), nSize                '将图片文件的字节复制到全局内存中
	            
	            lGSI.GdiplusVersion = 1
	            If GdiplusStartup(lToken, lGSI, 0) = 0 Then                     '初始化GDI+
	                If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then            '从全局内存创建流
	                    GdipLoadImageFromStream istm, lImage                    '将流中内容加载为GDI+ Image图形对象
	                    GdipCreateHBITMAPFromBitmap lImage, hBmp, &HFFFFFF      '从Image获取Bitmap句柄
	                    GdipDisposeImage lImage                                 '释放Image对象
	                    
	                    '以下代码从Bitmap句柄生成一个StdPicture对象
	                    CLSIDFromString StrPtr("{00020400-0000-0000-C000-000000000046}"), IID_IDispatch(0)
	                    With pic
	                        .Size = Len(pic)
	                        .Type = 1
	                        .hBmp = hBmp
	                        .hPal = 0
	                    End With
	                    OleCreatePictureIndirect pic, IID_IDispatch(0), 1, LoadWebImage
	                End If
	                GdiplusShutdown lToken '关闭GDI+
	            End If
	            GlobalUnlock hMem
	            GlobalFree hMem         '释放全局内存
	        End If
	    End With
	End Function



[http://club.excelhome.net/thread-400986-1-1.html](http://club.excelhome.net/thread-400986-1-1.html)

# 如何将Image控件中的图片数据转为二进制数组?(已解决!封装PropertyBag对象)  #

- 1.用GetDC获取一个窗口或桌面的DC句柄
- 2.用CreateCompatibleDC创建一个内存DC
- 3. 用SelectObject(hDC, Image1.Picture.Handle)将image1中的图片选入内存DC中，之前必须用Loadpicture或通过属性窗口向Image1中载入一幅图片
- 4. 再定义一个三维数组arrBits(0 to 3, lWidth-1, lHeight-1) as byte，这里的lWidth和lHeight是图形的宽和高，以像素为单位，stdPicture的长度单位是Himetric，需要乘以一个常量96 / 2540将其转换为像素
- 5.然后用GetDIBits将每个像素的RGB颜色值放到数组中

这样就将每个像素的RGB颜色信息放入arrBits数组中了，arrBits(0,x,y)表示从图片左下角算起横向第x,纵向第y个像素的蓝色亮度值，arrBits(1,x,y)和arrBits(2,x,y)则分别代表该点的绿色和红色的亮度，arrBits(3,x,y)为保留字节
	
	
代码：

	'函数功能：该函数检索一指定窗口的客户区域或整个屏幕的显示设备上下文环境的句柄，以后可以在GDI函数中使用该句柄来在设备上下文环境中绘图。
	hWnd：设备上下文环境被检索的窗口的句柄，如果该值为NULL，GetDC则检索整个屏幕的设备上下文环境。
	Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
	
	应用程序不能调用ReleaseDC函数来释放由CreateDC函数创建的设备上下文环境，只能使用DeleteDC函数。
	Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

	该函数创建一个与指定设备兼容的内存设备上下文环境（DC）。通过GetDc()获取的HDC直接与相关设备沟通，而本函数创建的DC，则是与内存中的一个表面相关联。
	Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
	Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
	
	该函数创建与指定的设备环境相关的设备兼容的位图。
	Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
	
	获取指定兼容位图的位，然后将其作一个DIB—设备无关位图（Device-Independent Bitmap）使用的指定格式复制到一个缓冲区中。	
	cScanLines：指定检索的扫描线数。
	lpvBits：指向用来检索位图数据的缓冲区的指针。如果此参数为NULL，那么函数将把位图的维数与格式传递给lpbi参数指向的BITMAPINFO结构。
	lpbi：指向一个BITMAPINFO结构的指针，此结构确定了设备所在位图的数据格式。
	uUsage：指定BITMAPINFO结构的bmiColors成员的格式
	Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitMapInfo, ByVal wUsage As Long) As Long

	该函数使用指定的DIB位图中发现的颜色数据来设置位图中的像素。
	hdc：指向设备环境中的句柄。
	hbmp：指向位图的句柄。函数要使用指定DIB中的颜色数据对该位图进行更改。
	uStartScan：为参数lpvBits指向的数组中的、与设备无关的颜色数据指定起始扫描线。
	cScanLines：为包含与设备无关的颜色数据的数组指定扫描线数目。
	lpvBits：指向DIB颜色数据的指针，这些数据存储在字节类型的数组中，位图值的格式取决于参数lpbmi指向的BITMAPINFO结构中的成员biBitCount。
	lpbmi：指向BITMAPINFO数据结构的指针，该结构包含有关DIB的信息。
	fuColorUse：指定是否提供了BITMAPINFO结构中的bmiColors成员，如果提供了，那么bmiColors是否包含了明确的RGB值或调色板索引
	Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitMapInfo, ByVal wUsage As Long) As Long
	Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
	
	Private Type BitMapInfoHeader ''文件信息头——BITMAPINFOHEADER
	    biSize As Long
	    biWidth As Long
	    biHeight As Long
	    biPlanes As Integer
	    biBitCount As Integer
	    biCompression As Long
	    biSizeImage As Long
	    biXPelsPerMeter As Long
	    biYPelsPerMeter As Long
	    biClrUsed As Long
	    biClrImportant As Long
	End Type
	
	Private Type RGBQuad
	    rgbBlue As Byte
	    rgbGreen As Byte
	    rgbRed As Byte
	    ''rgbReserved As Byte
	End Type
	
	Private Type BitMapInfo
	    bmiHeader As BitMapInfoHeader
	    bmiColors As RGBQuad
	End Type
	
	创建一个新的图片对象初始化根据到PICTDESC结构。
	Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
	lpDicDesc 一个包含图片状态的结构
	riid 接口标识符的引用，它描述了接口的类型。
	fOwn 如果为真，图片对象在对象被破坏时摧毁它的图片。如果为FALSE，调用者负责摧毁图片。
	lpUnk 接收riid参数的接口指针。当成功返回时，该参数包含了所需的新创建对象的接口指针。如果调用成功，调用者可以使用Release 函数来释放这个对象。如果调用失败，该参数值为NULL。

	Private Type GUID
	    Data1   As Long
	    Data2   As Integer
	    Data3   As Integer
	    Data4(7)   As Byte
	End Type
	Private Type PicBmp
	    Size   As Long
	    Type   As Long
	    hBmp   As Long
	    hPal   As Long
	    Reserved   As Long
	End Type
	
	Private Const HIMETRIC_PER_PIXEL = 96 / 2540
	Private Const vbPicTypeBitmap = 1
	
	Private Enum EnumPicMode
	    BlackWhite = 0
	    GrayScale = 1
	End Enum
		
	
	Private Sub Command1_Click()
	    Image2.Picture = Convert(Image1.Picture, GrayScale)
	End Sub
	
	Private Sub CommandButton1_Click()
	    Image3.Picture = Convert(Image1.Picture, BlackWhite, Slider1.Value)
	End Sub
	
	Private Sub Slider1_Click()
	    Label1.Caption = "阈值：" & Slider1.Value
	    Image3.Picture = Convert(Image1.Picture, BlackWhite, Slider1.Value)
	End Sub
	
	Private Sub UserForm_Initialize()
	    Label1.Caption = "阈值：" & Slider1.Value
	End Sub
	
	Private Function Convert(PicSrc As StdPicture, ToMode As EnumPicMode, Optional bytThreshold As Byte = 128) As StdPicture
	    Dim ix As Integer
	    Dim iy As Integer
	    Dim iWidth As Integer '以像素为单位的图形宽度
	    Dim iHeight As Integer '以像素为单位的图形高度
	    Dim bytTarget As Byte
	    Dim hDC As Long, hDCmem As Long
	    Dim hBmp As Long, hBmpPrev As Long
	    
	    Dim bits() As Byte '三维数组，用于获取原彩色图像中各像素的RGB数值以及存放转化后的灰度值
	    Dim bitsBW() As Byte '三维数组，用于存放转化为黑白图后各像素的值
	    
	    '获取图形的宽度和高度
	    iWidth = PicSrc.Width * HIMETRIC_PER_PIXEL
	    iHeight = PicSrc.Height * HIMETRIC_PER_PIXEL
	    
	    '创建并初始化一个bitMapInfo自定义类型
	    Dim bi24BitInfo As BitMapInfo
	    With bi24BitInfo.bmiHeader
	        .biBitCount = 32
	        .biCompression = 0&
	        .biPlanes = 1
	        .biSize = Len(bi24BitInfo.bmiHeader)
	        .biWidth = iWidth
	        .biHeight = iHeight
	    End With
	    '重新定义数组大小
	    ReDim bits(0 To 3, 0 To iWidth - 1, 0 To iHeight - 1) As Byte
	    hDC = GetDC(0)
	    hDCmem = CreateCompatibleDC(hDC)
	    '使用GetDIBits方法一次性获取picture1中各点的rgb值，比point方法或getPixel函数逐像素获取像素rgb要快出一个数量级
	    lrtn = GetDIBits(hDCmem, PicSrc.Handle, 0&, iHeight, bits(0, 0, 0), bi24BitInfo, DIB_RGB_COLORS)
	    '数组的三个维度分别代表像素的RGB分量、以图形左下角为原点的X和Y坐标。
	    '具体说来，这时bits(0,2,3)代表从图形左下角数起横向第2个纵向第3个像素的Blue值，而bits(1,2,3)和bits(2,2,3)分别的Green值和Red值.
	    
	   
	    If ToMode = GrayScale Then '***********RGB转为灰度******
	        For ix = 0 To iWidth - 1
	            For iy = 0 To iHeight - 1
	                'Debug.Print bits(0, ix, iy), bits(1, ix, iy), bits(2, ix, iy)
	                bytTarget = bits(0, ix, iy) * 0.11 + bits(1, ix, iy) * 0.59 + bits(2, ix, iy) * 0.3 '这是传统的根据三原色亮度加权得到灰阶的算法
	                bits(0, ix, iy) = bytTarget
	                bits(1, ix, iy) = bytTarget
	                bits(2, ix, iy) = bytTarget
	            Next
	        Next
	    Else '*********转为黑白图像********
	        For ix = 0 To iWidth - 1
	            For iy = 0 To iHeight - 1
	                bytTarget = bits(0, ix, iy) * 0.11 + bits(1, ix, iy) * 0.59 + bits(2, ix, iy) * 0.3
	                If bytTarget < bytThreshold Then
	                    bits(0, ix, iy) = 0
	                    bits(1, ix, iy) = 0
	                    bits(2, ix, iy) = 0
	                    Else
	                    bits(0, ix, iy) = 255
	                    bits(1, ix, iy) = 255
	                    bits(2, ix, iy) = 255
	                End If
	            Next
	        Next
	    End If
	    
	 '************下面是从DIBits转为stdPicture的代码***************
	    hBmp = CreateCompatibleBitmap(hDC, iWidth, iHeight) '创建一个与屏幕兼容的位图，得到它的句柄
	    SetDIBits hDCmem, hBmp, 0, iHeight, bits(0, 0, 0), bi24BitInfo, DIB_RGB_COLORS '将DIBits信息放入hBmp中
	    DeleteDC hDCmem
	    ReleaseDC 0, hDC
	    
	'从hBmp得到stdPicture的标准方法
	    Dim r     As Long
	    Dim pic     As PicBmp
	    Dim IPic     As StdPicture
	    Dim IID_IDispatch     As GUID
	    '填充IDispatch界面，clsID为{00020400-0000-0000-C000-000000000046}
	    With IID_IDispatch
	          .Data1 = &H20400
	          .Data4(0) = &HC0
	          .Data4(7) = &H46
	    End With
	    '填充Pic结构
	    With pic
	          .Size = Len(pic) 'pic结构的大小
	          .Type = vbPicTypeBitmap '图形类型, Bitmap
	          .hBmp = hBmp '位图句柄
	          .hPal = 0 '因为是24位色，所以不需要设定Pallete
	    End With
	    '建立Picture对象
	    r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
	    '返回Picture对象
	    Set Convert = IPic
	End Function
