Attribute VB_Name = "Kruskal"
Option Explicit

Type Edge
    begin As Long
    end As Long
    weight As Long
End Type

Sub MiniSpanTree_Kruskal()
    Dim i As Long, n As Long, m As Long
    Dim edges(14) As Edge, parent(8) As Long
    
    Range("E:G").Clear
    Range("E1").Value = "Lowcost"
    For i = 0 To 14
        edges(i).begin = Cells(i + 2, "b").Value
        edges(i).end = Cells(i + 2, "c").Value
        edges(i).weight = Cells(i + 2, "d").Value
    Next i
    
    For i = 0 To 9
        n = MyFind(parent, edges(i).begin)
        m = MyFind(parent, edges(i).end)
        If n <> m Then
            parent(n) = m
            Range("E65535").end(xlUp).Offset(1, 0) = edges(i).begin
            Range("F65535").end(xlUp).Offset(1, 0) = edges(i).end
            Range("G65535").end(xlUp).Offset(1, 0) = edges(i).weight
        End If
    Next i
    
End Sub

Function MyFind(parent() As Long, f As Long) As Long
    Do While parent(f) > 0
        f = parent(f)
    Loop
    MyFind = f
End Function
