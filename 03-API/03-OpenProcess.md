[http://club.excelhome.net/thread-1020309-1-1.html](http://club.excelhome.net/thread-1020309-1-1.html)

#关闭shell打开的文件#

	Sub aaa()
		Shell "C:\Program Files\winrar\winrar.exe", vbNormalFocus
	End Sub
	
	Sub ccc()
		Shell "taskkill /f /im winrar.exe"
	End Sub



	Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
	Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
	Dim hProcess As Long
	
	Private Sub Command1_Click() '打开进程
	    Dim pid As Long
	    pid = Shell("calc.exe", vbNormalFocus)
	    hProcess = OpenProcess(&H1, 0, pid)
	End Sub
	
	Private Sub Command2_Click() '关闭进程
	    TerminateProcess hProcess, 1
	End Sub