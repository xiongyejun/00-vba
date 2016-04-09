Attribute VB_Name = "ShSort"
Option Explicit

Const I_ShellSort As Long = 65535

Sub TestShellSort()

    Dim Arr(1 To I_ShellSort)
    Dim i As Long
    Dim t As Double

    For i = 1 To I_ShellSort
        Arr(i) = Cells(i, "b").Value
    Next i
    
    t = Timer
    Call ShellSort(Arr)
    Range("E7").Value = Timer - t
    Range("F7").Value = I_ShellSort & "¸öÊý¾Ý"
    
    SetShape "ShellSort", Range("D7")
    Range("C1:C65535").ClearContents
    Range("C1:C" & I_ShellSort).Value = Application.WorksheetFunction.Transpose(Arr)
    Erase Arr

End Sub
Sub ShellSort(l()) 'Ï£¶ûÅÅÐò
    Dim i As Long, j As Long
    Dim Increment As Long, Temp
    
    Increment = UBound(l)
    
    Do
        Increment = Increment \ 3 + 1
        
        For i = Increment + 1 To UBound(l)
            If l(i) < l(i - Increment) Then
                Temp = l(i)     'ÔÝ´æ
                
                j = i - Increment
                If j > 0 Then
                    Do While j > 0 And Temp < l(j)
                        l(j + Increment) = l(j)
                        j = j - Increment
                        If j <= 0 Then Exit Do
                    Loop
                End If
                l(j + Increment) = Temp
                
            End If
        
        Next i
    
    Loop While Increment > 1

End Sub
