Attribute VB_Name = "CountSort"
Option Explicit

Sub TestCountSort()
    Dim Arr(), i As Long
    Dim k As Long
    
    
    k = Range("B65536").End(xlUp).Row
    ReDim Arr(1 To k)
    
    For i = 1 To k
        Arr(i) = Cells(i, 2).Value
    Next i
    
    Range("F10").Value = k & "¸öÊý¾Ý"
    Call CountSort(Arr)
    Erase Arr
End Sub



Sub CountSort(Arr())
    Dim iMax As Long
    Dim Temp() As Long
    Dim ResultArr() As Long
    Dim i As Long
    Dim iCount As Long
    Dim t As Double
    t = Timer
    
    iCount = UBound(Arr)
    
    
    ReDim ResultArr(1 To iCount)
    iMax = Application.WorksheetFunction.Max(Arr)
    ReDim Temp(iMax) As Long
    
    For i = 1 To iCount
        Temp(Arr(i)) = Temp(Arr(i)) + 1
    Next i
    
    For i = 1 To iMax
        Temp(i) = Temp(i) + Temp(i - 1)
    Next i
    
    For i = iCount To 1 Step -1
        ResultArr(Temp(Arr(i))) = Arr(i)
        Temp(Arr(i)) = Temp(Arr(i)) - 1
    Next i
    
    Range("E10").Value = Timer - t
    
    Range("C1:C" & iCount).ClearContents
    Range("C1:C" & iCount).Value = Application.WorksheetFunction.Transpose(ResultArr)
    
    
    Erase Arr, Temp, ResultArr
End Sub
