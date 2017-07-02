Attribute VB_Name = "mkNAP"
Option Explicit

'背包问题
'1定空间下，有多重体积和价值的东西
'如果放才能最大价值
'自定向下的动态规划

Type Item
    Size As Long
    Val As Long
End Type

Type KnownItem
    Val As Long
    Index As Long   'items的下标
    Prev As Long '引用的上一个最优解
End Type

Sub TestKnap()
    Dim items() As Item
    Dim i As Long
    Dim maxKnown() As KnownItem
    Dim M As Long
    
    ReDim items(3) As Item
    i = 0
    items(i).Size = 3
    items(i).Val = 4
    i = i + 1
    
    items(i).Size = 4
    items(i).Val = 5
    i = i + 1
    
    items(i).Size = 7
    items(i).Val = 10
    i = i + 1
    
    items(i).Size = 8
    items(i).Val = 11
    i = i
    
    items(i).Size = 9
    items(i).Val = 13
    i = i

    M = 17
    ReDim maxKnown(M) As KnownItem
    For i = 0 To M
        maxKnown(i).Val = -1
        maxKnown(i).Index = -1
        maxKnown(i).Prev = -1
    Next
    For i = 0 To 2
        maxKnown(i).Val = 0
    Next
    maxKnown(i).Val = items(0).Val
    maxKnown(i).Index = 0
    
    Knap M, items, maxKnown
    
    For i = 0 To UBound(maxKnown)
        Cells(i + 2, 1).Value = i
        Cells(i + 2, 2).Value = maxKnown(i).Val
        If maxKnown(i).Index > -1 Then Cells(i + 2, 3).Value = items(maxKnown(i).Index).Val
        Cells(i + 2, 4).Value = maxKnown(i).Prev
    Next
    
    Dim str As String
    
    i = M
    str = maxKnown(i).Val & "="
    Do Until maxKnown(i).Prev = -1
        str = str & "+" & items(maxKnown(i).Index).Val & "(items(" & maxKnown(i).Index & "))"
        i = maxKnown(i).Prev
    Loop
    
    Range("E1").Value = str
End Sub

Function Knap(M As Long, items() As Item, maxKnown() As KnownItem) As Long
    Dim t  As Long
    Dim space As Long
    Dim max_val As Long
    Dim max_i As Long
    Dim max_Prev As Long
    Dim i As Long
    
    If maxKnown(M).Val > -1 Then
        Knap = maxKnown(M).Val
        Exit Function
    End If
    
    max_val = 0
    For i = 0 To UBound(items)
        space = M - items(i).Size
        If space >= 0 Then
            t = Knap(space, items, maxKnown) + items(i).Val
            If t > max_val Then
                max_val = t
                max_i = i
                max_Prev = space
            End If
        End If
    Next
    
    maxKnown(M).Val = max_val
    maxKnown(M).Index = max_i
    maxKnown(M).Prev = max_Prev
End Function






