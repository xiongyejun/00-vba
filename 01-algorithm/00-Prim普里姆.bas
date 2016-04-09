Attribute VB_Name = "Prim"
Option Explicit

Sub MiniSpanTree_Prim()
    Dim min As Long, i As Long, j As Long, k As Long
    Dim Adjvex(8) As Long    '������ض����±�
    Dim Lowcost(8) As Long   '������ض����ߵ�Ȩֵ
    Dim G()
    
    G = Range("B2:J10")
    Range("L:N").Clear
    Lowcost(0) = 0  '��ʼ����һ��ȨֵΪ0����V0����������
                    'lowcost��ֵΪ0����������Ǵ��±�Ķ����Ѿ�����������
    Adjvex(0) = 0   '��ʼ����һ�������±�Ϊ0
    
    For i = 1 To UBound(G, 1) - 1
        Lowcost(i) = G(1, i + 1)  '��V0������֮�бߵ�Ȩֵ��������
        Adjvex(i) = 0           '��ʼ����ΪV0���±�
    Next i
    
    For i = 1 To UBound(G, 1) - 1
        min = 65535
        j = 1: k = 0
        
        Do While (j < UBound(G, 1))
            If Lowcost(j) <> 0 And Lowcost(j) < min Then
                min = Lowcost(j)
                k = j
            End If
            j = j + 1
        Loop
        
        Range("L65535").end(xlUp).Offset(1, 0).Value = "V" & Adjvex(k)
        Range("M65535").end(xlUp).Offset(1, 0).Value = "V" & k
        Range("N65535").end(xlUp).Offset(1, 0).Value = min
'        MySleep
        
        Lowcost(k) = 0
        
        For j = 1 To UBound(G, 1) - 1
            If Lowcost(j) <> 0 And G(k + 1, j + 1) < Lowcost(j) Then
                Lowcost(j) = G(k + 1, j + 1)
                Adjvex(j) = k
            End If
        Next j
        
    Next i
    
'    Stop
End Sub

Private Sub MySleep()
    Dim t As Double
    t = Timer
    
    Do While t > Timer - 0.5
        DoEvents
    Loop
End Sub
