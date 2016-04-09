Attribute VB_Name = "线性查找"
Option Explicit

Sub TestFindMaxSum()
    Dim Arr() As Long, TempArr
    Dim i As Long
    Dim t As Double
    Dim MaxSum() As Long
    
    t = Timer
    
    TempArr = Range("A1:A" & Range("A65535").End(xlUp).Row).Value
    
    ReDim Arr(1 To Range("A65535").End(xlUp).Row)
    
    For i = 1 To Range("A65535").End(xlUp).Row
        Arr(i) = TempArr(i, 1)
    Next i
    
    
    MaxSum = FindMaxSum(Arr)
    Range("C6:E6").Value = MaxSum
    Range("f6").Value = Timer - t
    
    
    Erase TempArr
    Erase Arr
    Erase MaxSum
  
End Sub

Function FindMaxSum(Arr() As Long)
    Dim MaxSum(1 To 3) As Long
    Dim iSum As Long
    Dim iLeft As Long
    Dim iRight As Long
    Dim i As Long
    
    For i = 1 To UBound(Arr)
        If iSum >= 0 Then
            iSum = iSum + Arr(i)
            iRight = i
        Else
            iSum = Arr(i)           '如果小于0就抛弃前面的
            iLeft = i
            iRight = i
        End If
        
        If iSum > MaxSum(3) Then
            MaxSum(1) = iLeft
            MaxSum(2) = iRight
            MaxSum(3) = iSum
        End If
    
    Next i
    
    FindMaxSum = MaxSum
    
End Function
