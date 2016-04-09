Attribute VB_Name = "HeapSort"
Option Explicit

Sub TestHeap()
    Dim Arr() As Long, i As Long
    Dim k As Long, t As Double
    
'    k = Range("B65536").End(xlUp).Row
    k = 5000
    ReDim Arr(1 To k) As Long
    
    For i = 1 To k
        Arr(i) = Cells(i, 2).Value
    Next i
    
    t = Timer
    HeapSort Arr
    Range("E9").Value = Timer - t
    Range("F9").Value = k & "¸öÊý¾Ý"
    
    Range("C1:C" & k).Value = Application.WorksheetFunction.Transpose(Arr)
    
    Erase Arr
End Sub

Function HeapSort(Arr() As Long)
    Dim i As Long
    i = UBound(Arr)
    
    Do Until i = 1
        BulidMaxHeap Arr, i
        Swaplong Arr, 1, i
        i = i - 1
    Loop
    
End Function


Function MaxHeapify(Arr() As Long, i As Long, High As Long)
    Dim l As Long
    Dim r As Long
    Dim iMax As Long
    
    l = 2 * i ' Left(i)
    r = l + 1 ' Right(i)
    
    iMax = i
    If l <= High Then
        If Arr(l) > Arr(i) Then
            iMax = l
        End If
    End If
    
    If r <= High Then
        If Arr(r) > Arr(iMax) Then
            iMax = r
        End If
    End If
    
    If iMax <> i Then
        Swaplong Arr, i, iMax
        MaxHeapify Arr, iMax, High
    End If
    
End Function

Function BulidMaxHeap(Arr() As Long, High As Long)
    Dim i As Long
    
    For i = High / 2 To 1 Step -1
        MaxHeapify Arr, i, High
    Next i
End Function

Function Swaplong(l() As Long, Low As Long, High As Long)
    Dim iTemp As Long
    iTemp = l(Low)
    l(Low) = l(High)
    l(High) = iTemp
End Function
