Attribute VB_Name = "MMain"
Option Explicit

Sub TestSort2(control As IRibbonControl) 'SortDialog
    Dim arr()
    Dim i As Long, j As Long
    Const N As Long = 10
    Const COLS As Long = 5
    
    ReDim arr(1 To N, 1 To COLS)
    
    VBA.Randomize
    For i = 1 To N
        For j = 1 To COLS
            arr(i, j) = Int(VBA.Rnd() * N)
        Next j
    Next i
    
    Dim c_sorter As CSorter2
    
    Set c_sorter = New CSorter2
    c_sorter.Data = arr
    c_sorter.SortCol = 1
    
    Cells.Clear
    Range("a1").Resize(N, COLS).Value = arr
    Range("a1").Offset(0, COLS * 2 + 1).Value = N & "条数据用时" & MSort.QuickSort(c_sorter, 1, N)
    
    arr = c_sorter.Data
    Range("a1").Offset(0, COLS + 1).Resize(N, COLS).Value = arr
End Sub

Sub TestSort1(control As IRibbonControl) 'SortRemoveAllSorts
    Dim arr()
    Dim i As Long
    Const N As Long = 10
    
    ReDim arr(1 To N)
    
    VBA.Randomize
    For i = 1 To N
        arr(i) = Int(VBA.Rnd() * N)
    Next i
    
    Dim c_sorter As CSorter
    
    Set c_sorter = New CSorter
    c_sorter.Data = arr
    
    Cells.Clear
    Range("a1").Resize(N, 1).Value = Application.WorksheetFunction.Transpose(arr)
    Range("a1").Offset(0, 2).Value = N & "条数据用时" & MSort.QuickSort(c_sorter, 1, N)
    
    arr = c_sorter.Data
    Range("b1").Resize(N, 1).Value = Application.WorksheetFunction.Transpose(arr)
End Sub
