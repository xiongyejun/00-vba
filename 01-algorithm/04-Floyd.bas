Attribute VB_Name = "Floyd"
Option Explicit

Dim Patharc(8, 8) As Long '存储最短路径下标
Dim ShortPathTable(8, 8) As Long  '存储到各点最短路径的权值和

Sub Test()
    Dim arr(8, 8) As Long
    Dim i As Long, j As Long, Vx As Long
    
    For i = 2 To 10
        For j = 2 To 10
            arr(i - 2, j - 2) = Cells(i, j).Value
        Next j
    Next i
    
    ShortestPath_Floyd arr
    
    Range("B13:J21").Value = ShortPathTable
    Range("B24:J32").Value = Patharc
    
End Sub

Sub ShortestPath_Floyd(G() As Long)
    Dim v As Long, w As Long, k As Long
    Dim t As Double
    Range("B13:J21").Value = ""
    Range("B24:J32").Value = ""
    
    For v = 0 To UBound(G, 2)
        
        For w = 0 To UBound(G, 2)
            ShortPathTable(v, w) = G(v, w)
            Patharc(v, w) = w
        Next w
        
    Next v
    
    For k = 0 To UBound(G, 2)
    
        For v = 0 To UBound(G, 2)
            
            For w = 0 To UBound(G, 2)
                If ShortPathTable(v, w) > ShortPathTable(v, k) + ShortPathTable(k, w) Then
                    ShortPathTable(v, w) = ShortPathTable(v, k) + ShortPathTable(k, w)
                    Patharc(v, w) = Patharc(v, k)
'    t = Timer
'    Range("B13:J21").Value = ShortPathTable
'    Range("B24:J32").Value = Patharc
    
'   MySleep
                End If
            Next w
            
        Next v
        
    Next k
    MsgBox "完成"
End Sub
Private Sub MySleep()
    Dim t As Double
    t = Timer
    
    Do While t > Timer - 0.5
        DoEvents
    Loop
End Sub

Sub PrintV()

End Sub
