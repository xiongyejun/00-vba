[http://club.excelhome.net/thread-1430127-1-1.html](http://club.excelhome.net/thread-1430127-1-1.html)

# 已用区域  保存为图片 #

	Private Declare Function OpenClipboard Lib "User32" (ByVal hWnd As Long) As Long
	Private Declare Function CloseClipboard Lib "User32" () As Long
	Private Declare Function GetClipboardData Lib "User32" (ByVal uFormat As Long) As Long
	Private Declare Function CopyEnhMetaFileA Lib "Gdi32" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
	Private Declare Function DeleteEnhMetaFile Lib "Gdi32" (ByVal hdc As Long) As Long
	Sub lujkhua()
	    Dim picnm As String
	    picnm = Application.GetSaveAsFilename("try1.jpg", "图片, *.jpg", , "请选择保存路径并键入文件名")
	    ActiveSheet.UsedRange.CopyPicture
	    OpenClipboard 0
	    DeleteEnhMetaFile CopyEnhMetaFileA(GetClipboardData(14), picnm)
	    CloseClipboard
	End Sub