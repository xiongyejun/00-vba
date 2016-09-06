[http://club.excelhome.net/forum.php?mod=viewthread&tid=546023&extra=page%3D1](http://club.excelhome.net/forum.php?mod=viewthread&tid=546023&extra=page%3D1)


	Private Declare Function GetCursorPos Lib "user32"  (lpPoint As MyPoint) As Long
	Private Type MyPoint: X As Long: Y As Long: End Type
	Public flag

	Sub getshape()
	    Dim CurPos As MyPoint
	    Do While flag = True
	    GetCursorPos CurPos
	    x1 = CurPos.X: y1 = CurPos.Y
	    Set CurRng = ActiveWindow.RangeFromPoint(x1, y1)
	    If CurRng Is Nothing Then Exit Sub
	    On Error Resume Next
	    
	    If CurRng.ShapeRange.AutoShapeType = 138 Then
	        Range("A1") = CurRng.Name
	        For Each i In ActiveSheet.Shapes
	            If i.Fill.ForeColor.SchemeColor = 52 And i.Name <> CurRng.Name Then
	                i.Fill.ForeColor.SchemeColor = 65
	            End If
	        Next
	        CurRng.ShapeRange.Fill.ForeColor.SchemeColor = 52
	    End If
	    DoEvents
	    Loop
	End Sub
	
	Sub demo()
	    MsgBox "开始"
	    flag = True
	    getshape
	End Sub
	Sub enddemo()
	    MsgBox "结束"
	    flag = False
	End Sub


动态给shp指定宏


	Sub auto_add_macro()
	    Dim i As Long
	
	    '新建一个模型时手动运行，一次性添加宏
	     For i = 1 To ActiveSheet.Shapes.Count
	         '5表示对象类型是自选图形
	         If ActiveSheet.Shapes(i).Type = 5 Then
	            ActiveSheet.Shapes(i).OnAction = "'userclick(""" & ActiveSheet.Shapes(i).Name & """) '"
	         End If
	     Next
	
	End Sub
	
	Sub userclick(region_name)
	
	    Range("A1").Value = region_name
	'  ActiveSheet.Shapes(Range("A1").Value).Fill.ForeColor.SchemeColor = 9
	'
	''1、取A1单元格值，将上次选择的地图版块填充黄色，即还原填充色
	'
	'     Range("A1").Value = region_name
	'
	''2、将当前选择的地图版块名称填写到A1
	'
	'    ActiveSheet.Shapes(region_name).Fill.ForeColor.SchemeColor = 52
	
	    '3、将当前选择的地图版块填充红色
	    Range("h3").Value = region_name
	
	End Sub


hook钩子实现
	
	Option Explicit
	Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long '获取鼠标屏幕坐标
	'##############鼠标钩子相关的API函数及参数####################
	Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
	        ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long '设置钩子
	Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long '注消钩子，用完后一定要注销，以免影响WINDOWS速度
	Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, _
	        ByVal wParam As Long, lparam As Any) As Long '如果有其它钩子，则继续执行
	Private Const WH_MOUSE_LL As Long = 14
	Private Const WM_MOUSEMOVE = &H200
	'##############################################################
	Private Type POINTAPI
	    Mx As Long
	    My As Long
	End Type
	
	Public hHook As Long
	Public LastText As String
	
	Public Sub EnableHook() '设置钩子
	     If hHook = 0 Then
	        hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf HookProc, Application.Hinstance, 0)
	     End If
	End Sub
	
	Public Sub FreeHook() '注消钩子
	     If hHook <> 0 Then
	        Call UnhookWindowsHookEx(hHook)
	        hHook = 0 '使用了上面一句注消钩子后,还要加上这句才可以完全释放钩子
	     
	   '  Application.StatusBar = False
	     End If
	End Sub
	
	Public Function HookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lparam As Long) As Long '设置钩子函数,参数格式固定
	     
	Dim pt As POINTAPI, Rng As Range
	
	If nCode < 0 Then
	   HookProc = CallNextHookEx(hHook, nCode, wParam, lparam) '如果有其它钩子，则继续执行,不然要出错
	   Exit Function
	End If
	     
	If wParam = WM_MOUSEMOVE Then '如果为鼠标移动事件,执行以下语句
	        '#######以下为您要赋予钩子的语句,因为要写入WINDOW,故不能调用其它子过程#####
	   Call GetCursorPos(pt)
	   On Error Resume Next
	   ActiveSheet.Shapes("biaoqian").TextFrame.Characters.Text = ""
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	   Dim objShape
	   Dim AltText As String
	   
	   Set objShape = ActiveWindow.RangeFromPoint(x:=pt.Mx, y:=pt.My)
	         
	         If objShape.Type = 5 Then
	            
	            AltText = objShape.Name
	            Set objShape = Nothing
	            
	            If ActiveSheet.Shapes(AltText).Type = 5 Then
	               ActiveSheet.Shapes(LastText).Fill.ForeColor.SchemeColor = 9
	               ActiveSheet.Shapes(AltText).Fill.ForeColor.SchemeColor = 52
	               Cells(1, 1) = AltText
	               LastText = AltText
	            End If
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	           ' Application.StatusBar = "Mouse Cursor At " & CStr(pt.Mx) & "," & CStr(pt.My) & "  " & AltText
	        
	            With Sheet1.Shapes("biaoqian")
	               .TextFrame.Characters.Text = Application.VLookup(Cells(1, 1).Text, Sheet5.[a:e], 2, False)
	               .Top = Pixel2PointY(pt.My - ActiveWindow.PointsToScreenPixelsY(Cells(1, 1).Top)) + 10 ' Pixel2PointY
	               .Left = Pixel2PointX(pt.Mx - ActiveWindow.PointsToScreenPixelsX(Cells(1, 1).Left)) + 10
	            End With
	'##########################################################################
	         Else
	            
	            AltText = ""
	        
	         End If
	 
	End If
	
	End Function
	
	
