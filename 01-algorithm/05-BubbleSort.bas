Attribute VB_Name = "BSort"
Option Explicit
Const I_Bubble As Long = 3000

Sub TestBubbleSort0()
    Dim Arr(1 To I_Bubble)
    Dim i As Long
    Dim t As Double
    
    For i = 1 To I_Bubble
        Arr(i) = Cells(i, "b").Value
    Next i
    
    t = Timer
    Call BubbleSort0(Arr)
    Range("E3").Value = Timer - t
    Range("F3").Value = I_Bubble & "个数据"
    
    SetShape "Bsort0", Range("D3")
    Range("C1:C65535").ClearContents
    Range("C1:C" & I_Bubble).Value = Application.WorksheetFunction.Transpose(Arr)
    Erase Arr

End Sub

Sub BubbleSort0(l())
    Dim i As Long, j As Long
    
    For i = 1 To UBound(l)
        For j = i + 1 To UBound(l)
            If l(i) > l(j) Then
                Call Swap(l, i, j)
            End If
        Next j
    Next i
End Sub

Sub TestBubbleSort()
    Dim Arr(1 To I_Bubble)
    Dim i As Long
    Dim t As Double
    
    For i = 1 To I_Bubble
        Arr(i) = Cells(i, "b").Value
    Next i
    
    t = Timer
    Call BubbleSort(Arr)
    Range("E4").Value = Timer - t
    Range("F4").Value = I_Bubble & "个数据"
    
    SetShape "Bsort", Range("D4")
    Range("C1:C65535").ClearContents
    Range("C1:C" & I_Bubble).Value = Application.WorksheetFunction.Transpose(Arr)
    Erase Arr

End Sub
Sub BubbleSort(l())
    Dim i As Long, j As Long
    
    For i = 1 To UBound(l)
        For j = UBound(l) - 1 To i Step -1
            If l(j) > l(j + 1) Then
                Call Swap(l, j, j + 1)
            End If
        Next j
    Next i
End Sub

Sub TestBubbleSort2()
    Dim Arr(1 To I_Bubble)
    Dim i As Long
    Dim t As Double
    
    For i = 1 To I_Bubble
        Arr(i) = Cells(i, "b").Value
    Next i
    
    t = Timer
    Call BubbleSort2(Arr)
    Range("E5").Value = Timer - t
    Range("F5").Value = I_Bubble & "个数据"
    
    SetShape "Bsort2", Range("D5")
    Range("C1:C65535").ClearContents
    Range("C1:C" & I_Bubble).Value = Application.WorksheetFunction.Transpose(Arr)
    Erase Arr

End Sub

Sub BubbleSort2(l())
    Dim i As Long, j As Long
    Dim Flag As Boolean
    
    For i = 1 To UBound(l)
        Flag = False
        For j = UBound(l) - 1 To i Step -1
            If l(j) > l(j + 1) Then
                Call Swap(l, j, j + 1)
                Flag = True
            End If
        Next j
        If Not Flag Then Exit For
    Next i
End Sub
