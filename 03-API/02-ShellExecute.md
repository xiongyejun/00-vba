[http://club.excelhome.net/forum.php?mod=viewthread&tid=1200080&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=1200080&extra=page%3D1)

	'API函数声明
	Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long	
	
	Sub open我的电脑()
	    ShellExecute Application.hwnd, "open", "explorer", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", 0, 1
	End Sub
	
	Sub open网上邻居()
	    ShellExecute Application.hwnd, "open", "explorer", "::{208d2c60-3aea-1069-a2d7-08002b30309d}", 0, 1
	End Sub
	
	Sub open回收站()
	    ShellExecute Application.hwnd, "open", "explorer", "::{645ff040-5081-101b-9f08-00aa002f954e}", 0, 1
	End Sub
	
	Sub open控制面板()
	    ShellExecute Application.hwnd, "open", "explorer", "::{21ec2020-3aea-1069-a2dd-08002b30309d}", 0, 1
	End Sub
	
	Sub open我的文档()
	    ShellExecute Application.hwnd, "open", "explorer", 0, 0, 1
	End Sub
	
	Sub open文件夹()
	    ShellExecute Application.hwnd, "open", "C:\Program Files\Tencent\QQ\Bin", 0, 0, 1
	End Sub
	
	Sub open应用程序()
	    ShellExecute Application.hwnd, "open", "C:\Program Files\Tencent\QQ\Bin\QQ.exe", 0, 0, 1
	End Sub
	
	Sub openExcelFile()
	    ShellExecute Application.hwnd, "open", "C:\text.xls", 0, 0, 1
	End Sub
	
	Sub openTextFile()
	    ShellExecute Application.hwnd, "open", "C:\text.txt", 0, 0, 1
	End Sub
	
	Sub open网页()
	    ShellExecute Application.hwnd, "open", "http://club.excelhome.net", 0, 0, 1
	End Sub


用API做这个有点大材小用啦。

vba的shell方法同样可以实现的：

	Sub open我的电脑()
	    Shell "explorer ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
	End Sub
	Sub open网上邻居()
	    Shell "explorer ::{208d2c60-3aea-1069-a2d7-08002b30309d}"
	End Sub
	Sub open回收站()
	    Shell "explorer ::{645ff040-5081-101b-9f08-00aa002f954e}"
	End Sub
	Sub open控制面板()
	    Shell "explorer ::{21ec2020-3aea-1069-a2dd-08002b30309d}"
	End Sub
	Sub open我的文档()
	    Shell "explorer"
	End Sub
	Sub open文件夹()
	    Shell "explorer C:\Program Files\Tencent\QQ\Bin"
	End Sub
	Sub open应用程序()
	    Shell "C:\Program Files\Tencent\QQ\Bin\QQ.exe"
	End Sub
	Sub openExcelFile()
	    Shell "excel D:\test\1.xls"
	End Sub
	Sub openTextFile()
	    Shell "notepad D:\test\1.txt"
	End Sub
	Sub open网页()
	    Shell "explorer http://club.excelhome.net"
	End Sub