Attribute VB_Name = "MMain"
Option Explicit

Sub vba_main()
    Dim c As CStack
    Dim i As Long
    Dim node_val As Long

    Set c = New CStack
    VBA.Randomize
    Range("A:b").Clear
    Dim n As CNode
    For i = 1 To 100000
        node_val = VBA.Rnd() * 100000

'        Cells(i, 1).Value = node_val
        
        Set n = New CNode
        n.Value = node_val
        c.Push n
        
'        If i Mod 2 Then
'            Set n = c.Pop()
''            Range("B" & VBA.CStr(Cells.Rows.Count)).End(xlUp).Offset(1, 0).Value = "pop " & n.Value
'            Set n = Nothing
'        End If
        
        'VBA.DoEvents
    Next i
    
'    Set n = c.Pop()
'    Do Until n Is Nothing
''        Range("B" & VBA.CStr(Cells.Rows.Count)).End(xlUp).Offset(1, 0).Value = "pop " & n.Value
'        Set n = Nothing
'        Set n = c.Pop()
'    Loop


    Range("A1").Select

    Set n = Nothing
    Set c = Nothing
End Sub

