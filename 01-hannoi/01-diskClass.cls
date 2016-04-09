VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "diskClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim myShp As Shape
Dim myLevel As Long

Public Property Get Level() As Long        '层
    Level = myLevel
End Property
Public Property Let Level(x As Long)       '层
    myLevel = x
End Property

Public Property Get Shape() As Shape         '获取图形
    Set Shape = myShp
End Property

Public Property Set Shape(shp As Shape)      '图形
    Set myShp = shp
End Property

Function setShapePosition(iLeft As Long, iTop As Long)
    Me.Shape.Left = iLeft
    Me.Shape.Top = iTop
End Function

Function diskUp(iCol As Long)
'    Me.Shape.Top = Pillar.Top
    
    moveVertical -10, Pillar.Top
    diskUpCol = iCol + 1
End Function

Function diskDown(iCol As Long)
    Dim iTop As Long, iLeft As Long
    
    iLeft = iCol * 200 + 100 - Me.Shape.Width / 2
    
    
    If stack(0, iCol) = 0 Then
        iTop = Pillar.Top + Pillar.Height - DISK_HEIGHT / 2
    Else
        iTop = diskArr(stack(stack(0, iCol), iCol)).Shape.Top
    End If
    
    setShapePosition iLeft, Me.Shape.Top
'    moveHorizontal 10, iLeft
'    setShapePosition Me.Shape.Left, iTop - DISK_HEIGHT / 2
    moveVertical 10, iTop - DISK_HEIGHT / 2
    
    stack(stack(0, iCol) + 1, iCol) = diskUpNumber      '栈的push过程
    stack(0, iCol) = stack(0, iCol) + 1
    diskUpCol = 0
    
    iSteps = iSteps + 1
    Range("I1").Value = iSteps
    
    If stack(0, 2) = diskNumber Then
        MsgBox "完成任务。", vbInformation, "hannoi"
        endGame
    End If
End Function

Private Function moveVertical(i As Long, iTop As Long)
    Dim sign As Long
    
    If i > 0 Then
        sign = 1
    Else
        sign = -1
    End If
    
    ifSleep = True
    Do Until sign * Me.Shape.Top >= sign * iTop
        If Not ifSleep Then Exit Function
        Me.Shape.Top = Me.Shape.Top + i
        mySleep 0.01
    Loop
    
    Me.Shape.Top = iTop
    ifSleep = False
End Function

Private Function moveHorizontal(i As Long, iLeft As Long)
    Dim sign As Long
    
    If i > 0 Then
        sign = 1
    Else
        sign = -1
    End If
    
    ifSleep = True
    Do Until sign * Me.Shape.Left >= sign * iLeft
        Me.Shape.Left = Me.Shape.Left + i
        mySleep 0.01
    Loop
    
    Me.Shape.Left = iLeft
    ifSleep = False
End Function

