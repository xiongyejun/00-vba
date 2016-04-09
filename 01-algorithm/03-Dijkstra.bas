Attribute VB_Name = "Dijkstra"
Option Explicit

Dim Patharc(8) As Long  '�洢���·���±�
Dim ShortPathTable(8) As Long   '�洢���������·����Ȩֵ��

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

'����ͼG��V0���㵽���ඥ��V���·��P(V)����Ȩ����D(V)
Function ShortestPath_Dijkstra(G() As Long, V0 As Long)
    Dim v As Long
    Dim w As Long
    Dim k As Long
    Dim min As Long
     
    Dim Final(8) As Long    'final(w)=1��ʾ���V0��Vw�����·��
    
    For v = 0 To UBound(G, 2)
        Final(v) = 0    'ȫ�������ʼ��Ϊδ֪���·��״̬
        ShortPathTable(v) = G(V0, v) '����V0�������ߵĶ������Ȩֵ
        Patharc(v) = 0        '��ʼ��·������PΪ0
    Next v
    
    ShortPathTable(V0) = 0           'V0��V0·��Ϊ0
    Final(V0) = 1       'V0ֻV0����Ҫ��·��
    
    '��ʼ��ѭ����ÿ�����V0��ÿ��V��������·��
    For v = 1 To UBound(G, 2) - 1
        min = 65535
        
        For w = 0 To UBound(G, 2)
            If Final(w) = 0 And ShortPathTable(w) <> 0 And ShortPathTable(w) < min Then
                k = w
                min = ShortPathTable(w)
            End If
        Next w
        
        Final(k) = 1    '��Ŀǰ�ҵ�������Ķ�����Ϊ1
        
        For w = 0 To UBound(G, 2)   '������ǰ���·��������
            '�������V�����·������������·���ĳ��ȶ̵Ļ�
            If Final(w) = 0 And (min + G(k, w) < ShortPathTable(w)) Then
                '˵���ҵ��˸��̵�·�����޸�D(w)��P(w)
                ShortPathTable(w) = min + G(k, w)
                Patharc(w) = k
            End If
        Next w
        
    Next v
    
    Range("L2").Resize(9, 1).Value = Application.WorksheetFunction.Transpose(Patharc)
    Range("m2").Resize(9, 1).Value = Application.WorksheetFunction.Transpose(ShortPathTable)
End Function

