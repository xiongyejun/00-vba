Attribute VB_Name = "模块1"
Option Explicit
Public Const MIN_DISK_WIDTH As Long = 50      '最小的盘子的宽
Public Const DISK_STEP As Long = 10           '盘子宽度增长
Public Const DISK_HEIGHT As Long = 50         '盘子的厚

Public Enum Pillar                            '柱子图片的宽、高、左、上
    Width = 600
    Height = 320
    Left = 0
    Top = 40
End Enum

Public diskArr() As New diskClass            'disk的数组
Public diskNumber As Long                    'disk的总数量，就是层数
Public colorArr(1 To 10) As Long
Public diskUpCol As Long                     'disk在上面第几列,0代表没有,
Public diskUpNumber As Long                  '在上面的是几号盘子

Public stack(10, 2) As Long                  '栈
Public iSteps As Long                        '已经操作的步数
Public ifSleep As Boolean                    '是否在sleep过程中

Sub main()
    Dim iTop As Long
    Const iWidth As Long = 100
    Const iHeight As Long = 40
    Const iStep As Long = 200
    Dim iLeft As Long
    
    On Error GoTo errHandle
    
    ifSleep = False
    diskUpCol = 0      '没有disk在上面
    diskUpNumber = 0
    
    stack(0, 0) = 0     '空栈
    stack(0, 1) = 0     '空栈
    stack(0, 2) = 0     '空栈
    iSteps = 0
    Range("I1").Value = 0
    
    ActiveWindow.DisplayGridlines = False   '不显示网格线
    deleteShape                             '删除椭圆
    diskNumber = Range("C1").Value          '盘子的总数量
    If diskNumber > 10 Then
        MsgBox "盘子数量太多了，会很累的，不玩了。", vbInformation, "hannoi"
        endGame
        Exit Sub
    End If
    Range("E1").Value = 2 ^ diskNumber - 1
    
    ReDim diskArr(1 To diskNumber) As New diskClass
    
    setColor
    addShape
    
    setShape "pillar", Pillar.Left, Pillar.Top, Pillar.Width, Pillar.Height
    ActiveSheet.Shapes("pillar").ZOrder msoSendToBack
    
    iTop = Pillar.Top + Pillar.Height - 5
    iLeft = Pillar.Left + iStep / 2 - iWidth / 2
    
    setShape "A", iLeft, iTop, iWidth, iHeight: iLeft = iLeft + iStep
    setShape "B", iLeft, iTop, iWidth, iHeight: iLeft = iLeft + iStep
    setShape "C", iLeft, iTop, iWidth, iHeight
    
    Exit Sub
    
errHandle:
    MsgBox Err.Description
    endGame
End Sub

Sub endGame()
    ifSleep = False
    Application.DisplayFullScreen = False
    
    Erase diskArr
    Erase stack
    Erase colorArr
End Sub

Sub macroA()
    macro 0
End Sub
Sub macroB()
    macro 1
End Sub
Sub macroC()
    macro 2
End Sub

Function macro(iCol As Long)
    On Error GoTo Err
    
    If ifSleep Then Exit Function
    
    If diskUpCol = 0 Then    '没有disk在上面
        diskUpNumber = stack(stack(0, iCol), iCol)
        
        If diskUpNumber <> 0 Then
            diskArr(diskUpNumber).diskUp (iCol)
            stack(stack(0, iCol), iCol) = 0         'up的disk，变为0——栈的pop过程
            stack(0, iCol) = stack(0, iCol) - 1     '栈的哨兵-1
        End If
    Else        '有disk在上面
        If stack(0, iCol) = 0 Then
            diskArr(diskUpNumber).diskDown (iCol)
        ElseIf diskArr(diskUpNumber).Level < diskArr(stack(stack(0, iCol), iCol)).Level Then
            diskArr(diskUpNumber).diskDown (iCol)
        End If
        
    End If
    
Err:

End Function

Private Function addShape()
    Dim shp As Shape, i As Long
    Dim iLeft As Long, iTop As Long, iWidth As Long, iHeight As Long, iColor As Long, strText As String
    
    iTop = Pillar.Top + Pillar.Height - DISK_HEIGHT / 2
    
    For i = diskNumber To 1 Step -1
        iWidth = MIN_DISK_WIDTH + i * DISK_STEP
        iLeft = 100 - iWidth / 2
        iTop = iTop - DISK_HEIGHT / 2
        
        Set shp = ActiveSheet.Shapes.addShape(msoShapeOval, iLeft, iTop, iWidth, DISK_HEIGHT)
        With shp
            .Fill.ForeColor.SchemeColor = colorArr(i)
            .Line.Visible = msoFalse
'            .TextFrame.Characters.Text = i
            .Name = "disk" & i
'            .TextFrame.HorizontalAlignment = xlCenter
        End With
        
        Set diskArr(i).Shape = shp
        diskArr(i).Level = i
        
        stack(diskNumber - i + 1, 0) = i
        stack(0, 0) = stack(0, 0) + 1
    Next i
    
    Set shp = Nothing
End Function

Function setColor()
    Dim i As Long
    Const strColor As String = "6、 4、 38、 54、 37、 48、 10、 34、 46、 7"
    
    For i = 1 To 10
        colorArr(i) = Cells(i + 1, "m").Interior.ColorIndex
        Cells(i + 1, "n").Value = colorArr(i)
        If colorArr(i) = -4142 Then colorArr(i) = Split(strColor, "、")(i - 1)
    Next i
    
End Function

Sub deleteShape()
    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        If shp.Type = 1 Then
            shp.Delete
        End If
    Next shp
    Set shp = Nothing
End Sub

Sub setShape(shpName As String, iLeft As Long, iTop As Long, iWidth As Long, iHeight As Long)
    With ActiveSheet.Shapes(shpName)
        .LockAspectRatio = msoFalse '取消锁定纵横比
        .Left = iLeft
        .Top = iTop
        .Width = iWidth ' Rng.Width
        .Height = iHeight 'Rng.Height
    End With
End Sub

Sub baiduBaiKe()
    Dim strMsg As String
    
    strMsg = "汉诺塔：又称河内塔，问题是源于印度一个古老传说的益智玩具。"
    strMsg = strMsg & vbNewLine & vbNewLine & "大梵天创造世界的时候做了三根金刚石柱子，在一根柱子上从下往上按照大小顺序摞着64片黄金圆盘。"
    strMsg = strMsg & vbNewLine & vbNewLine & "大梵天命令婆罗门把圆盘从下面开始按大小顺序重新摆放在另一根柱子上。"
    strMsg = strMsg & vbNewLine & vbNewLine & "并且规定，在小圆盘上不能放大圆盘，在三根柱子之间一次只能移动一个圆盘。"

    MsgBox strMsg, vbInformation, "hannoi"
End Sub

Sub getHanoiStep()
    Dim stepArr() As String
    Dim n As Long
    Dim k As Long, i As Long
    
    Randomize
    n = Int(Rnd * 20) + 1
    If Application.InputBox("请问" & n & "层汉诺塔至少需要多少步完成转移？", Title:="hannoi", Type:=1) <> 2 ^ n - 1 Then
        MsgBox "回答错误！", vbExclamation, "hannoi"
        Exit Sub
    End If
    
    k = 1
    n = Range("C1").Value
    If n > 10 Then
        MsgBox "盘子数量太多了，会很累的，不玩了。", vbInformation, "hannoi"
        Exit Sub
    End If
    
    ReDim stepArr(1 To 2 ^ n - 1) As String
    hanoi n, "A", "B", "C", stepArr, k
    
    Range("L:L").ClearContents
    
    Range("L2").Resize(2 ^ n - 1, 1).Value = Application.WorksheetFunction.Transpose(stepArr)
    
    If MsgBox("是否看一下单元格C1数量的动画演示？", vbYesNo, "hannoi") <> vbYes Then
        Exit Sub
    End If
    
    Application.DisplayFullScreen = True
    Call main
    For i = 1 To k - 1
        Range("L:L").Interior.ColorIndex = 0
        Range("L" & i + 1).Interior.ColorIndex = 6
        Range("K1").Value = stepArr(i)
        
'        callMacro Mid(stepArr(i), 1, 1)
        callMacro VBA.Left(stepArr(i), 1)
'        callMacro Mid(stepArr(i), 3, 1)
        callMacro VBA.Right(stepArr(i), 1)
    Next i
    Range("K1").Value = "完成"
    Range("L:L").ClearContents
    Range("L:L").Interior.ColorIndex = 0
    
    Application.DisplayFullScreen = False
    
    Erase stepArr
End Sub

Private Function callMacro(str As String)
    Select Case str
        Case "A"
            Call macroA
        Case "B"
            Call macroB
        Case "C"
            Call macroC
    End Select
End Function

Sub hanoi(n As Long, from As String, denpend_on As String, des As String, stepArr() As String, k As Long)
        
    If n = 1 Then
        stepArr(k) = from & "→" & des
        k = k + 1
    Else
        hanoi n - 1, from, des, denpend_on, stepArr, k
        stepArr(k) = from & "→" & des
        k = k + 1
        hanoi n - 1, denpend_on, from, des, stepArr, k
    End If
End Sub

Function mySleep(i As Single)
    Dim t As Single
    t = Timer
    
    Do Until Timer - t > i
        DoEvents
    Loop
End Function
