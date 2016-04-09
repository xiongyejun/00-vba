Attribute VB_Name = "ģ��1"
Option Explicit
Public Const MIN_DISK_WIDTH As Long = 50      '��С�����ӵĿ�
Public Const DISK_STEP As Long = 10           '���ӿ������
Public Const DISK_HEIGHT As Long = 50         '���ӵĺ�

Public Enum Pillar                            '����ͼƬ�Ŀ��ߡ�����
    Width = 600
    Height = 320
    Left = 0
    Top = 40
End Enum

Public diskArr() As New diskClass            'disk������
Public diskNumber As Long                    'disk�������������ǲ���
Public colorArr(1 To 10) As Long
Public diskUpCol As Long                     'disk������ڼ���,0����û��,
Public diskUpNumber As Long                  '��������Ǽ�������

Public stack(10, 2) As Long                  'ջ
Public iSteps As Long                        '�Ѿ������Ĳ���
Public ifSleep As Boolean                    '�Ƿ���sleep������

Sub main()
    Dim iTop As Long
    Const iWidth As Long = 100
    Const iHeight As Long = 40
    Const iStep As Long = 200
    Dim iLeft As Long
    
    On Error GoTo errHandle
    
    ifSleep = False
    diskUpCol = 0      'û��disk������
    diskUpNumber = 0
    
    stack(0, 0) = 0     '��ջ
    stack(0, 1) = 0     '��ջ
    stack(0, 2) = 0     '��ջ
    iSteps = 0
    Range("I1").Value = 0
    
    ActiveWindow.DisplayGridlines = False   '����ʾ������
    deleteShape                             'ɾ����Բ
    diskNumber = Range("C1").Value          '���ӵ�������
    If diskNumber > 10 Then
        MsgBox "��������̫���ˣ�����۵ģ������ˡ�", vbInformation, "hannoi"
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
    
    If diskUpCol = 0 Then    'û��disk������
        diskUpNumber = stack(stack(0, iCol), iCol)
        
        If diskUpNumber <> 0 Then
            diskArr(diskUpNumber).diskUp (iCol)
            stack(stack(0, iCol), iCol) = 0         'up��disk����Ϊ0����ջ��pop����
            stack(0, iCol) = stack(0, iCol) - 1     'ջ���ڱ�-1
        End If
    Else        '��disk������
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
    Const strColor As String = "6�� 4�� 38�� 54�� 37�� 48�� 10�� 34�� 46�� 7"
    
    For i = 1 To 10
        colorArr(i) = Cells(i + 1, "m").Interior.ColorIndex
        Cells(i + 1, "n").Value = colorArr(i)
        If colorArr(i) = -4142 Then colorArr(i) = Split(strColor, "��")(i - 1)
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
        .LockAspectRatio = msoFalse 'ȡ�������ݺ��
        .Left = iLeft
        .Top = iTop
        .Width = iWidth ' Rng.Width
        .Height = iHeight 'Rng.Height
    End With
End Sub

Sub baiduBaiKe()
    Dim strMsg As String
    
    strMsg = "��ŵ�����ֳƺ�������������Դ��ӡ��һ�����ϴ�˵��������ߡ�"
    strMsg = strMsg & vbNewLine & vbNewLine & "�����촴�������ʱ�������������ʯ���ӣ���һ�������ϴ������ϰ��մ�С˳������64Ƭ�ƽ�Բ�̡�"
    strMsg = strMsg & vbNewLine & vbNewLine & "���������������Ű�Բ�̴����濪ʼ����С˳�����°ڷ�����һ�������ϡ�"
    strMsg = strMsg & vbNewLine & vbNewLine & "���ҹ涨����СԲ���ϲ��ܷŴ�Բ�̣�����������֮��һ��ֻ���ƶ�һ��Բ�̡�"

    MsgBox strMsg, vbInformation, "hannoi"
End Sub

Sub getHanoiStep()
    Dim stepArr() As String
    Dim n As Long
    Dim k As Long, i As Long
    
    Randomize
    n = Int(Rnd * 20) + 1
    If Application.InputBox("����" & n & "�㺺ŵ��������Ҫ���ٲ����ת�ƣ�", Title:="hannoi", Type:=1) <> 2 ^ n - 1 Then
        MsgBox "�ش����", vbExclamation, "hannoi"
        Exit Sub
    End If
    
    k = 1
    n = Range("C1").Value
    If n > 10 Then
        MsgBox "��������̫���ˣ�����۵ģ������ˡ�", vbInformation, "hannoi"
        Exit Sub
    End If
    
    ReDim stepArr(1 To 2 ^ n - 1) As String
    hanoi n, "A", "B", "C", stepArr, k
    
    Range("L:L").ClearContents
    
    Range("L2").Resize(2 ^ n - 1, 1).Value = Application.WorksheetFunction.Transpose(stepArr)
    
    If MsgBox("�Ƿ�һ�µ�Ԫ��C1�����Ķ�����ʾ��", vbYesNo, "hannoi") <> vbYes Then
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
    Range("K1").Value = "���"
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
        stepArr(k) = from & "��" & des
        k = k + 1
    Else
        hanoi n - 1, from, des, denpend_on, stepArr, k
        stepArr(k) = from & "��" & des
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
