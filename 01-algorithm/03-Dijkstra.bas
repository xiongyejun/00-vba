Attribute VB_Name = "Dijkstra"
Option Explicit

Dim Patharc(8) As Long  '存储最短路径下标
Dim ShortPathTable(8) As Long   '存储到各点最短路径的权值和

Sub Test()
    Dim arr(8, 8) As Long
    Dim i As Long, j As Long, Vx As Long
    
    For i = 2 To 10
        For j = 2 To 10
            arr(i - 2, j - 2) = Cells(i, j).Value
        Next j
    Next i
    
    Vx = 0
    ShortestPath_Dijkstra arr, Vx
    
    Range("n2:o10").Clear
    i = Range("N1").Value
    Range("n2").Value = "V" & i
    j = 3
    
    Do While i <> Vx
        Cells(j, "O").Value = arr(i, Patharc(i))
        i = Patharc(i)
        Cells(j, "N").Value = "V" & i
        j = j + 1
    Loop
    
End Sub

'有向图G的V0顶点到其余顶点V最短路径P(V)及带权长度D(V)
Function ShortestPath_Dijkstra(G() As Long, V0 As Long)
    Dim v As Long
    Dim w As Long
    Dim k As Long
    Dim min As Long
     
    Dim Final(8) As Long    'final(w)=1表示求得V0至Vw的最短路径
    
    For v = 0 To UBound(G, 2)
        Final(v) = 0    '全部顶点初始化为未知最短路径状态
        ShortPathTable(v) = G(V0, v) '将与V0店有连线的顶点加上权值
        Patharc(v) = 0        '初始化路径数组P为0
    Next v
    
    ShortPathTable(V0) = 0           'V0至V0路径为0
    Final(V0) = 1       'V0只V0不需要求路径
    
    '开始主循环，每次求得V0到每个V顶点的最短路径
    For v = 1 To UBound(G, 2) - 1
        min = 65535
        
        For w = 0 To UBound(G, 2)
            If Final(w) = 0 And ShortPathTable(w) <> 0 And ShortPathTable(w) < min Then
                k = w
                min = ShortPathTable(w)
            End If
        Next w
        
        Final(k) = 1    '将目前找到的最近的顶点置为1
        
        For w = 0 To UBound(G, 2)   '修正当前最短路径及距离
            '如果经过V顶点的路径比现在这条路径的长度短的话
            If Final(w) = 0 And (min + G(k, w) < ShortPathTable(w)) Then
                '说明找到了更短的路径，修改D(w)和P(w)
                ShortPathTable(w) = min + G(k, w)
                Patharc(w) = k
            End If
        Next w
        
    Next v
    
    Range("L2").Resize(9, 1).Value = Application.WorksheetFunction.Transpose(Patharc)
    Range("m2").Resize(9, 1).Value = Application.WorksheetFunction.Transpose(ShortPathTable)
End Function

