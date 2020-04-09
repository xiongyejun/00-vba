Attribute VB_Name = "MTestTree"
Option Explicit

Type Point
    x As Long
    y As Long
End Type

Sub vba_main()
    Dim ctree As CAVLTree
    Dim i As Long
    Dim node_val As Long

    Range("A:B").Clear
    Set ctree = New CAVLTree
    VBA.Randomize
    For i = 1 To 2 ^ 9 - 1
        node_val = VBA.Rnd() * 100
        Cells(i, 1).Value = node_val
        ctree.Add node_val
    Next i
    ctree.Balance
    
    Cells(1, 2).Value = ctree.SelectValue(111)
    
'    DeleteShp
'    ctree.DrawTree

    Range("A1").Select

    Set ctree = Nothing
End Sub

Function DrawOval(i_left As Long, i_top As Long, NodeValue As Long)
    Dim shp As Shape

    Set shp = ActiveSheet.Shapes.AddShape(msoShapeOval, i_left, i_top, 18, 17.25)
    shp.Fill.Visible = msoFalse

    With shp.TextFrame2.TextRange
        .Characters.Text = NodeValue
        shp.Select
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorText1

        With shp.TextFrame2
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
       End With
    End With
    MySleep 0.5
End Function

Function getChildNode(parentPoint As Point, iWeight As Double, ByVal leftAngel As Long, ByVal rightAngel As Long)
    Dim arr(1, 1) As Long
    Dim c As Double
    Dim b As Double
    Dim a As Double
    Const pi As Double = 3.1415926

    c = iWeight
    a = c * Sin(leftAngel / 180 * pi)   'Y
    b = c * Cos(leftAngel / 180 * pi)   'X

    arr(1, 0) = parentPoint.x - b
    arr(1, 1) = parentPoint.y + a

    a = c * Math.Sin(rightAngel / 180 * pi)   'Y
    b = c * Cos(rightAngel / 180 * pi)   'X

    arr(0, 0) = parentPoint.x - b
    arr(0, 1) = parentPoint.y + a

    getChildNode = arr
End Function

Function drawLine(x1 As Long, y1 As Long, x2 As Long, y2 As Long, iWidth As Double)
    Dim shp As Shape

    Set shp = ActiveSheet.Shapes.AddLine(x1, y1, x2, y2)
    With shp
        .Line.Weight = iWidth
        .Line.ForeColor.RGB = RGB(0, 204, 0)
        .Placement = xlFreeFloating
    End With

End Function

Function DeleteShp()
    Dim shp As Shape

    For Each shp In ActiveSheet.Shapes '
        If shp.Type <> 8 Then shp.Delete
    Next
End Function

Function MySleep(t As Double)
'    Dim t_now As Double
'
'    t_now = Timer
'
'    Do While Timer - t < t_now
'        DoEvents
'    Loop
End Function
