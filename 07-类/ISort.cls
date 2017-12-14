VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------交换------------
Function Swap(i As Long, j As Long)

End Function

'---------赋值------------
'将第j个赋值给第i个
Function Assignment(i As Long, j As Long)

End Function
'将Value赋值给第i个
Function AssignmentValue(i As Long, Value)

End Function
'将第i个赋值给Value
Function ReAssignmentValue(i As Long, Value)

End Function

'---------比较------------
'第i个是否小于第j个
Function Less(i As Long, j As Long) As Boolean

End Function
'第i个是否小于Value
Function LessValue(i As Long, Value) As Boolean

End Function
'第i个是否大于Value
Function LagerValue(i As Long, Value) As Boolean

End Function

'---------其他------------
'三数取中
Function MedianOfThree(ByVal Low As Long, ByVal High As Long)

End Function

'Attribute VB_Name = "MSort"
'Option Explicit
'
''返回用时（秒）
'Function QuickSort(I_Sort As ISort, Low As Long, High As Long) As Double  '快速排序
'    Dim Pivot As Long
'    Dim t As Double
'
'    t = Timer
'    If High - Low > 50 Then
'        Do While Low < High
'            Pivot = MyPartition(I_Sort, Low, High)
'            QuickSort I_Sort, Low, Pivot - 1     '对低子表递归排序
'            Low = Pivot + 1                       '尾递归
'        Loop
'    Else
'        InsertSort I_Sort, Low, High
'    End If
'
'    QuickSort = Timer - t
'End Function
'
'Private Function MyPartition(I_Sort As ISort, ByVal Low As Long, ByVal High As Long) As Long
'    Dim PivotKey        '枢轴
'
'    PivotKey = I_Sort.MedianOfThree(Low, High) '三数取中
'    Do While Low < High
'        Do While Low < High And (Not I_Sort.LessValue(High, PivotKey))
'            High = High - 1
'        Loop
'        I_Sort.Assignment Low, High
'
'        Do While Low < High And I_Sort.LessValue(Low, PivotKey)
'            Low = Low + 1
'        Loop
'        I_Sort.Assignment High, Low
'    Loop
'
'    I_Sort.AssignmentValue Low, PivotKey
'
'    MyPartition = Low
'End Function
'
''插入排序
'Private Function InsertSort(I_Sort As ISort, Low As Long, High As Long)
'    Dim i As Long, j As Long
'    Dim ShaoBing
'
'    For i = Low + 1 To High
'        If I_Sort.Less(i, i - 1) Then
'            I_Sort.ReAssignmentValue i, ShaoBing  '设置哨兵
'
'            j = i - 1
'            Do While I_Sort.LagerValue(j, ShaoBing)
'                I_Sort.Assignment j + 1, j
'                j = j - 1
'                If j = Low - 1 Then Exit Do
'            Loop
'
'            I_Sort.AssignmentValue j + 1, ShaoBing
'        End If
'    Next i
'End Function
'


'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "CSorter"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Option Explicit
''一维数组实现接口
'
'Implements ISort
'
'Private arr() As Variant
'
'Property Let Data(Value As Variant)
'    arr = Value
'End Property
'Property Get Data() As Variant
'    Data = arr
'End Property
'
''比较
'Private Function ISort_Less(i As Long, j As Long) As Boolean
'    ISort_Less = arr(i) < arr(j)
'End Function
'Private Function ISort_LessValue(i As Long, Value As Variant) As Boolean
'    ISort_LessValue = arr(i) < Value
'End Function
'Private Function ISort_LagerValue(i As Long, Value As Variant) As Boolean
'    ISort_LagerValue = arr(i) > Value
'End Function
'
''交换
'Private Function ISort_Swap(i As Long, j As Long) As Variant
'    Dim tmp As Variant
'
'    tmp = arr(i)
'    arr(i) = arr(j)
'    arr(j) = tmp
'End Function
'
''赋值
'Private Function ISort_Assignment(i As Long, j As Long) As Variant
'    arr(i) = arr(j)
'End Function
'Private Function ISort_AssignmentValue(i As Long, Value As Variant) As Variant
'    arr(i) = Value
'End Function
'Private Function ISort_ReAssignmentValue(i As Long, Value As Variant) As Variant
'    Value = arr(i)
'End Function
'
''三数取中
'Private Function ISort_MedianOfThree(ByVal Low As Long, ByVal High As Long) As Variant
'    Dim m As Long
'
'    m = Low + (High - Low) / 2
'
'    If arr(Low) > arr(High) Then Call ISort_Swap(Low, High)  '交换左端与右端数据，保证左端较小
'    If arr(m) > arr(High) Then Call ISort_Swap(High, m)      '交换中间与右端数据，保证中间较小
'    If arr(m) > arr(Low) Then Call ISort_Swap(m, Low)        '交换中间与左端数据，保证左端为中间值
'    ISort_MedianOfThree = arr(Low)
'End Function




'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "CSorter2"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Option Explicit
''二维数组实现接口
'
'Implements ISort
'
'Private Type LowHigh
'    Low As Long
'    High As Long
'End Type
'
'Private arr() As Variant
'Private i_col As Long           '要排序的列
'Private Col_Bound As LowHigh    '列的上下标
'
'Property Let SortCol(iCol As Long)
'    i_col = iCol
'End Property
'
'Property Let Data(Value As Variant)
'    arr = Value
'
'    Col_Bound.Low = LBound(arr, 2)
'    Col_Bound.High = UBound(arr, 2)
'End Property
'Property Get Data() As Variant
'    Data = arr
'End Property
'
''比较
'Private Function ISort_Less(i As Long, j As Long) As Boolean
'    ISort_Less = arr(i, i_col) < arr(j, i_col)
'End Function
'Private Function ISort_LessValue(i As Long, Value As Variant) As Boolean
'    ISort_LessValue = arr(i, i_col) < Value(i_col)
'End Function
'Private Function ISort_LagerValue(i As Long, Value As Variant) As Boolean
'    ISort_LagerValue = arr(i, i_col) > Value(i_col)
'End Function
'
''交换
'Private Function ISort_Swap(i As Long, j As Long) As Variant
'    Dim tmp As Variant
'    Dim k As Long
'
'    For k = Col_Bound.Low To Col_Bound.High
'        tmp = arr(i, k)
'        arr(i, k) = arr(j, k)
'        arr(j, k) = tmp
'    Next k
'End Function
'
''赋值
'Private Function ISort_Assignment(i As Long, j As Long) As Variant
'    Dim k As Long
'
'    For k = Col_Bound.Low To Col_Bound.High
'        arr(i, k) = arr(j, k)
'    Next k
'End Function
'Private Function ISort_AssignmentValue(i As Long, Value As Variant) As Variant
'    Dim k As Long
'
'    For k = Col_Bound.Low To Col_Bound.High
'        arr(i, k) = Value(k)
'    Next k
'End Function
'Private Function ISort_ReAssignmentValue(i As Long, Value As Variant) As Variant
'    Dim k As Long
'
'    ReDim Value(Col_Bound.Low To Col_Bound.High) As Variant
'    For k = Col_Bound.Low To Col_Bound.High
'        Value(k) = arr(i, k)
'    Next k
'End Function
'
''三数取中
'Private Function ISort_MedianOfThree(ByVal Low As Long, ByVal High As Long) As Variant
'    Dim m As Long
'    Dim Value() As Variant '返回某一行的1维数组
'    Dim k As Long
'
'    m = Low + (High - Low) / 2
'
'    If arr(Low, i_col) > arr(High, i_col) Then Call ISort_Swap(Low, High)  '交换左端与右端数据，保证左端较小
'    If arr(m, i_col) > arr(High, i_col) Then Call ISort_Swap(High, m)      '交换中间与右端数据，保证中间较小
'    If arr(m, i_col) > arr(Low, i_col) Then Call ISort_Swap(m, Low)        '交换中间与左端数据，保证左端为中间值
'
'    ISort_ReAssignmentValue Low, Value
'
'    ISort_MedianOfThree = Value
'End Function




'Attribute VB_Name = "MMain"
'Option Explicit
'
'Sub TestSort2(control As IRibbonControl) 'SortDialog
'    Dim arr()
'    Dim i As Long, j As Long
'    Const N As Long = 10
'    Const COLS As Long = 5
'
'    ReDim arr(1 To N, 1 To COLS)
'
'    VBA.Randomize
'    For i = 1 To N
'        For j = 1 To COLS
'            arr(i, j) = Int(VBA.Rnd() * N)
'        Next j
'    Next i
'
'    Dim c_sorter As CSorter2
'
'    Set c_sorter = New CSorter2
'    c_sorter.Data = arr
'    c_sorter.SortCol = 1
'
'    Cells.Clear
'    Range("a1").Resize(N, COLS).Value = arr
'    Range("a1").Offset(0, COLS * 2 + 1).Value = N & "条数据用时" & MSort.QuickSort(c_sorter, 1, N)
'
'    arr = c_sorter.Data
'    Range("a1").Offset(0, COLS + 1).Resize(N, COLS).Value = arr
'End Sub
'
'Sub TestSort1(control As IRibbonControl) 'SortRemoveAllSorts
'    Dim arr()
'    Dim i As Long
'    Const N As Long = 10
'
'    ReDim arr(1 To N)
'
'    VBA.Randomize
'    For i = 1 To N
'        arr(i) = Int(VBA.Rnd() * N)
'    Next i
'
'    Dim c_sorter As CSorter
'
'    Set c_sorter = New CSorter
'    c_sorter.Data = arr
'
'    Cells.Clear
'    Range("a1").Resize(N, 1).Value = Application.WorksheetFunction.Transpose(arr)
'    Range("a1").Offset(0, 2).Value = N & "条数据用时" & MSort.QuickSort(c_sorter, 1, N)
'
'    arr = c_sorter.Data
'    Range("b1").Resize(N, 1).Value = Application.WorksheetFunction.Transpose(arr)
'End Sub

