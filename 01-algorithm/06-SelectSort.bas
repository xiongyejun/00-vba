Attribute VB_Name = "SSort"
Option Explicit

Const I_SelectSort As Long = 6000

Sub TestSelectSort()

    Dim Arr(1 To I_SelectSort)
    Dim i As Long
    Dim t As Double

    For i = 1 To I_SelectSort
        Arr(i) = Cells(i, "b").Value
    Next i
    
    t = Timer
    Call SelectSort(Arr)
    Range("E6").Value = Timer - t
    Range("F6").Value = I_SelectSort & "¸öÊý¾Ý"
    
    SetShape "SelectSort", Range("D6")
    Range("C1:C65535").ClearContents
    Range("C1:C" & I_SelectSort).Value = Application.WorksheetFunction.Transpose(Arr)
    Erase Arr

End Sub

Sub SelectSort(l())
    Dim i As Long, j As Long, min As Long
    
    For i = 1 To UBound(l)
        min = i
        
        For j = i + 1 To UBound(l)
            If l(min) > l(j) Then
                min = j
            End If
        Next j
        
        If i <> min Then Call Swap(l, i, min)
   
    Next i
End Sub
