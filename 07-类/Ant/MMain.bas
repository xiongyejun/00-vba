Attribute VB_Name = "MMain"
Option Explicit

Sub vba_main()
    Dim tsp As CTSP
    
    Set tsp = New CTSP
    tsp.Distance = ReadTSP("E:\00-学习资料\2018年02月05日\蚂蚁\ulysses16.tsp")
    tsp.ItCount = 100
    tsp.Go
    
End Sub

Function ReadTSP(fileName As String) As Double()
    Dim num_file As Integer
    Dim str As String
    
    num_file = VBA.FreeFile
    
    Open fileName For Input As #num_file
    Line Input #num_file, str
    Dim iCount As Long
    iCount = VBA.Val(VBA.Replace(str, "NAME: ulysses", ""))
    
    Dim i As Long
    For i = 1 To 6
        Line Input #num_file, str
    Next
    
    Dim arrPoint() As Double
    ReDim arrPoint(iCount - 1, 1) As Double
    Dim tmp
    For i = 0 To iCount - 1
        Line Input #num_file, str
        tmp = VBA.Split(str, " ")
        arrPoint(i, 0) = tmp(2) 'X坐标
        arrPoint(i, 1) = tmp(3) 'Y坐标
    Next
    Close #num_file
    
    Dim Arr() As Double, j As Long
    ReDim Arr(iCount - 1, iCount - 1) As Double
    For i = 0 To iCount - 1
        For j = i + 1 To iCount - 1
            Arr(i, j) = CountDistance(arrPoint(i, 0), arrPoint(i, 1), arrPoint(j, 0), arrPoint(j, 1))
            Arr(j, i) = Arr(i, j)
        Next
    Next
    
    ReadTSP = Arr
End Function

Function CountDistance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
   CountDistance = ((y2 - y1) ^ 2 + (x2 - x1) ^ 2) ^ 0.5
End Function
