VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim iTemp As Long
    
    If Not ifStart Then Exit Sub        '游戏是否开始
    If Target.Cells.Count > 1 Then Exit Sub
    If Application.Intersect(rngGame, Target) Is Nothing Then Exit Sub
    
    Application.EnableEvents = False    '关闭事件、防止代码的select引发
    
    If selectRng Is Nothing Then
        selectChange Target
    ElseIf Application.Intersect(selectRng, Target) Is Nothing Then '选中的单元格不是再次选中
        selectChange Target
    Else                                                           '选中的单元格是再次选中，合并单元格
        With Target
            iTemp = selectRng.Range("A1").value
            getScore                    '得分
            selectRng.value = ""
            .value = iTemp + 1
            If myMax < iTemp + 1 Then   '获取最大值
                myMax = iTemp + 1
                rngMax.value = myMax
            End If
            
            If myMax < Application.WorksheetFunction.Max(rngGame) Then
                MsgBox "你手动改了数字！~~~"
                Call endGame
                Exit Sub
            End If
            
            Set selectRng = Nothing
            Range("A1").Select
            moveRng                     '填补空白
            
            If gameOver Then            '挂了
                MsgBox "你挂了！"
                PaiMing
                Call endGame
            End If
        End With
    End If
    
    Application.EnableEvents = True
End Sub

Function selectChange(Target As Range)
    Set selectRng = Target
     
    selectSameRng Target, 0, -1
    selectSameRng Target, 0, 1
    selectSameRng Target, 1, 0
    selectSameRng Target, -1, 0
    
    selectRng.Select
End Function

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
End Sub

Function getScore()
    rngScore.value = rngScore.value + selectRng.Cells.Count * selectRng.Range("A1").value
End Function
