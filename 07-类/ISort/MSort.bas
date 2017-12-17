Attribute VB_Name = "MSort"
Option Explicit

'返回用时（秒）
Function QuickSort(I_Sort As ISort, Low As Long, High As Long) As Double  '快速排序
    Dim Pivot As Long
    Dim t As Double
    
    t = Timer
    If High - Low > 50 Then
        Do While Low < High
            Pivot = MyPartition(I_Sort, Low, High)
            QuickSort I_Sort, Low, Pivot - 1     '对低子表递归排序
            Low = Pivot + 1                       '尾递归
        Loop
    Else
        InsertSort I_Sort, Low, High
    End If
    
    QuickSort = Timer - t
End Function

Private Function MyPartition(I_Sort As ISort, ByVal Low As Long, ByVal High As Long) As Long
    Dim PivotKey        '枢轴
    
    PivotKey = I_Sort.MedianOfThree(Low, High) '三数取中
    Do While Low < High
        Do While Low < High And (Not I_Sort.LessValue(High, PivotKey))
            High = High - 1
        Loop
        I_Sort.Assignment Low, High
        
        Do While Low < High And I_Sort.LessValue(Low, PivotKey)
            Low = Low + 1
        Loop
        I_Sort.Assignment High, Low
    Loop
    
    I_Sort.AssignmentValue Low, PivotKey
    
    MyPartition = Low
End Function

'插入排序
Private Function InsertSort(I_Sort As ISort, Low As Long, High As Long)
    Dim i As Long, j As Long
    Dim ShaoBing
    
    For i = Low + 1 To High
        If I_Sort.Less(i, i - 1) Then
            I_Sort.ReAssignmentValue i, ShaoBing  '设置哨兵
                    
            j = i - 1
            Do While I_Sort.LagerValue(j, ShaoBing)
                I_Sort.Assignment j + 1, j
                j = j - 1
                If j = Low - 1 Then Exit Do
            Loop
            
            I_Sort.AssignmentValue j + 1, ShaoBing
        End If
    Next i
End Function

