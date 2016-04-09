Attribute VB_Name = "ISort"
Option Explicit

Sub TestInsertSort()
    Dim Arr(), i As Long
    Dim k As Long, t As Double
    
'    ReDim Arr(1 To 5)
'    Arr(1) = 5
'    Arr(2) = 3
'    Arr(3) = 4
'    Arr(4) = 6
'    Arr(5) = 2
'    Call InsertSort(Arr, 1, UBound(Arr))

    k = 6000
    ReDim Arr(1 To k)

    For i = 1 To k
        Arr(i) = Cells(i, 2).Value
    Next i

    t = Timer

    Call InsertSort(Arr, 1, UBound(Arr))
    Range("E2").Value = Timer - t
    Range("F2").Value = UBound(Arr) & "个数据"

    SetShape "InsertSort", Range("D2")
    Range("C1:C" & k).ClearContents
    Range("C1:C" & k).Value = Application.WorksheetFunction.Transpose(Arr)
    Erase Arr
End Sub


Sub InsertSort(l(), Low As Long, High As Long)
    Dim i As Long, j As Long
    Dim ShaoBing
    
    For i = Low + 1 To High
    
        If l(i) < l(i - 1) Then
            ShaoBing = l(i)             '设置哨兵
                    
            j = i - 1
            Do While l(j) > ShaoBing
                l(j + 1) = l(j)
                j = j - 1
                If j = Low - 1 Then Exit Do
            Loop
            
            l(j + 1) = ShaoBing
        End If
    
    Next i
End Sub
