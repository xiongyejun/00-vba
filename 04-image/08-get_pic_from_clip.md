[http://club.excelhome.net/thread-997609-3-1.html](http://club.excelhome.net/thread-997609-3-1.html)

这个程序说起来也不复杂，涉及到两类API函数，一类是剪贴板函数，用于读取剪贴板中的位图句柄，另外一类是GDI+函数，用于位图到jpg文件的转换，具体代码如下：

	Option Explicit
	Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
	Private Declare Function CloseClipboard Lib "user32" () As Long
	Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
	Private Const CF_BITMAP = 2
	
	Private Type GUID
		 Data1 As Long
		 Data2 As Integer
		 Data3 As Integer
		 Data4(0 To 7) As Byte
	End Type
	
	Private Type GdiplusStartupInput
		 GdiplusVersion As Long
		 DebugEventCallback As Long
		 SuppressBackgroundThread As Long
		 SuppressExternalCodecs As Long
	End Type
	
	Private Type EncoderParameter
		 GUID As GUID
		 NumberOfValues As Long
		 type As Long
	 Value As Long
	End Type
	
	Private Type EncoderParameters
		 Count As Long
		 Parameter As EncoderParameter
	End Type
	
	Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
	使用GDI+ API之前，必须先调用GdiplusStartup这个函数，作用是初始化GDI+函数库。
	第一个参数是指向一个32位的无符号整型的指针，也就是指向一个汇编中的DWORD变量的指针，用于接受GDI+的TOKEN.TOKEN可以暂时理解成一个句柄，就像窗口的句柄类似。这个参数在调用GdiplusShutdown的时候用到。这个函数在结束GDI+编程后调用，起作用是释放GDI+的资源。 
	第二个以及第三个参数是指向两个结构体变量的指针。

	Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
	该函数会清理由Microsoft Windows GDI+使用过的资源。每次调用GdiplusStartup函数都要对应的使用一次该函数来完成清理工作。

	Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long


	Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
	'利用lBitmap该值保存图像

	Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
	
	Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
	从句柄创建 GDI+ 图像(lBitmap)
	
	Sub test()
	    Select Case CliptoJPG("c:\test.jpg")
	        Case 0:
	            MsgBox "剪贴板图片已保存"
	        Case 1:
	            MsgBox "剪贴板图片保存失败"
	        Case 2:
	            MsgBox "剪贴板中无图片"
	        Case 3:
	            MsgBox "剪贴板无法打开，可能被其他程序所占用"
	    End Select
	End Sub
	
	
	Private Function CliptoJPG(ByVal destfilename As String, Optional ByVal quality As Byte = 80) As Integer
	'*****该函数用于取出剪贴板中图片转换为jpg文件另存到指定路径****
	'参数说明：
	'     destfilename:要保存的jpg文件的完整路径，必要参数；
	'     quality: jpg文件的质量，0-100之间的数值，数值越大，图片质量越高
	'返回值：
	'     0-保存成功；1-保存失败；2-剪贴板中无位图数据；3-无法打开剪贴板
	
	    Dim tSI As GdiplusStartupInput
	    Dim lRes As Long
	    Dim lGDIP As Long
	    Dim lBitmap As Long
	    Dim hBmp As Long
	    
	    '尝试打开剪贴板
	    If OpenClipboard(0) Then
	        '尝试取出剪贴板中位图的句柄
	        hBmp = GetClipboardData(CF_BITMAP)
	        '如果hBmp为0，说明剪贴板中没有存放图片
	        If hBmp = 0 Then
	            CliptoJPG = 2
	            CloseClipboard
	            Exit Function
	        End If
	        CloseClipboard
	    Else   '如果openclipboard返回0(False)，说明剪贴板被其他程序所占用
	        CliptoJPG = 3
	        Exit Function
	    End If
	    
	    '初始化 GDI+
	    tSI.GdiplusVersion = 1
	    lRes = GdiplusStartup(lGDIP, tSI, 0)
	     
	    If lRes = 0 Then
	        '从句柄创建 GDI+ 图像
	        lRes = GdipCreateBitmapFromHBITMAP(hBmp, 0, lBitmap)
	         
	        If lRes = 0 Then
	            Dim tJpgEncoder As GUID
	            Dim tParams As EncoderParameters
	             
	            '初始化解码器的GUID标识
	            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
	             
	            '设置解码器参数
	            tParams.Count = 1
	            With tParams.Parameter ' Quality
	                '得到Quality参数的GUID标识
	                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
	                .NumberOfValues = 1
	                .type = 4
	                .Value = VarPtr(quality)
	            End With
	             
	            '保存图像
	            lRes = GdipSaveImageToFile(lBitmap, StrPtr(destfilename), tJpgEncoder, tParams)
	            If lRes = 0 Then
	                CliptoJPG = 0  '转换成功
	            Else
	                CliptoJPG = 1  '转换失败
	            End If
	             
	            '销毁GDI+图像
	            GdipDisposeImage lBitmap
	        End If
	         
	        '销毁 GDI+
	        GdiplusShutdown lGDIP
	    End If
	End Function