'Levenshtein算法
Function Compare(d As DataStruct) As Double
    Dim iLen1 As Long, iLen2 As Long
    Dim arr() As Long
    
    iLen1 = VBA.Len(d.Str1)
    iLen2 = VBA.Len(d.Str2)
    
    If (iLen1 * iLen2) = 0 Then
        Compare = 0
        Exit Function
    End If
    
    ReDim arr(iLen1, iLen2) As Long
    Dim i As Long, j As Long
    '初始化第1列
    For i = 0 To iLen1
        arr(i, 0) = i
    Next
    '初始化第1行
    For i = 0 To iLen2
        arr(0, i) = i
    Next
    
    Dim tmp As Long
    For i = 1 To iLen2
        For j = 1 To iLen1
            If VBA.Mid$(d.Str1, j, 1) = VBA.Mid$(d.Str2, i, 1) Then
                tmp = 0
            Else
                tmp = 1
            End If
            '等于左上角
            arr(j, i) = arr(j - 1, i - 1) + tmp
            '左
            If arr(j, i) > arr(j - 1, i) + 1 Then arr(j, i) = arr(j - 1, i) + 1
            '上
            If arr(j, i) > arr(j, i - 1) + 1 Then arr(j, i) = arr(j, i - 1) + 1
        Next j
    Next i
'    Range("B2").Resize(iLen1 + 1, iLen2 + 1).Value = arr
    tmp = iLen1
    If tmp < iLen2 Then tmp = iLen2
    Compare = 1 - arr(iLen1, iLen2) / tmp
End Function