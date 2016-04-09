Attribute VB_Name = "分治策略1"
Option Explicit

'三种情况
'1、在左边
'2、在中间——左边+中点+右边都包含
'3、在右边


Sub TestFindMaxCrossingSubArray1()
    Dim Arr() As Long, TempArr
    Dim i As Long
    Dim MaxSum() As Long
    Dim t As Double
    
    t = Timer
    TempArr = Range("A1:A" & Range("A65535").End(xlUp).Row).Value
    
    ReDim Arr(1 To Range("A65535").End(xlUp).Row)
    
    For i = 1 To Range("A65535").End(xlUp).Row
        Arr(i) = TempArr(i, 1)
    Next i
    
    MaxSum = FindMaximumSubArray1(Arr, 1, Range("A65535").End(xlUp).Row)

    
    Range("C5:E5").Value = MaxSum
    Range("f5").Value = Timer - t
    
    Erase TempArr
    Erase Arr
    Erase MaxSum
End Sub



Function FindMaxCrossingSubArray1(Arr() As Long, ByVal Low As Long, ByVal Mid As Long, ByVal High As Long)
    Dim LeftSum(1 To 3) As Long, RightSum(1 To 3) As Long
    Dim Sum As Long, i As Long, j As Long
    Dim MaxLeft As Long, MaxRight As Long
    Dim MaxSum(1 To 3) As Long
    
    LeftSum(3) = -9999999
    Sum = 0
    
    For i = Mid To Low Step -1
        Sum = Sum + Arr(i)
        
        If Sum > LeftSum(3) Then
            LeftSum(3) = Sum
            MaxLeft = i
        End If
    Next i
    
    Sum = 0
    RightSum(3) = -9999999
    For j = Mid + 1 To High
        Sum = Sum + Arr(j)
        If Sum > RightSum(3) Then
            RightSum(3) = Sum
            MaxRight = j
        End If
    Next j
    
'    ReDim MaxSum(1 To 3) As Long
    MaxSum(3) = LeftSum(3) + RightSum(3)
    MaxSum(1) = MaxLeft
    MaxSum(2) = MaxRight
    
    FindMaxCrossingSubArray1 = MaxSum

    Erase LeftSum
    Erase RightSum
    Erase MaxSum
End Function

Function FindMaximumSubArray1(Arr() As Long, ByVal Low As Long, ByVal High As Long)
    Dim Mid As Long
    Dim LeftSum() As Long
    Dim RightSum() As Long
    Dim CrossSum() As Long
    Dim TempSum(1 To 3) As Long
    
    If High = Low Then
        TempSum(3) = Arr(Low)
        TempSum(1) = Low
        TempSum(2) = Low
        FindMaximumSubArray1 = TempSum
    Else
        Mid = (Low + High) \ 2
        LeftSum = FindMaximumSubArray1(Arr(), Low, Mid)
        RightSum = FindMaximumSubArray1(Arr(), Mid + 1, High)
        CrossSum = FindMaxCrossingSubArray1(Arr(), Low, Mid, High)
        
        If LeftSum(3) >= RightSum(3) And LeftSum(3) > CrossSum(3) Then
            FindMaximumSubArray1 = LeftSum
        ElseIf RightSum(3) >= LeftSum(3) And RightSum(3) > CrossSum(3) Then
            FindMaximumSubArray1 = RightSum

        Else
            FindMaximumSubArray1 = CrossSum

        End If
        
    End If
    
    Erase LeftSum
    Erase RightSum
    Erase CrossSum
    Erase TempSum
End Function

