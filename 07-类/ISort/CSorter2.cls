VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSorter2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'二维数组实现接口

Implements ISort

Private Type LowHigh
    Low As Long
    High As Long
End Type

Private arr() As Variant
Private i_col As Long           '要排序的列
Private Col_Bound As LowHigh    '列的上下标

Property Let SortCol(iCol As Long)
    i_col = iCol
End Property

Property Let Data(Value As Variant)
    arr = Value
    
    Col_Bound.Low = LBound(arr, 2)
    Col_Bound.High = UBound(arr, 2)
End Property
Property Get Data() As Variant
    Data = arr
End Property

'比较
Private Function ISort_Less(i As Long, j As Long) As Boolean
    ISort_Less = arr(i, i_col) < arr(j, i_col)
End Function
Private Function ISort_LessValue(i As Long, Value As Variant) As Boolean
    ISort_LessValue = arr(i, i_col) < Value(i_col)
End Function
Private Function ISort_LagerValue(i As Long, Value As Variant) As Boolean
    ISort_LagerValue = arr(i, i_col) > Value(i_col)
End Function

'交换
Private Function ISort_Swap(i As Long, j As Long) As Variant
    Dim tmp As Variant
    Dim k As Long
    
    For k = Col_Bound.Low To Col_Bound.High
        tmp = arr(i, k)
        arr(i, k) = arr(j, k)
        arr(j, k) = tmp
    Next k
End Function

'赋值
Private Function ISort_Assignment(i As Long, j As Long) As Variant
    Dim k As Long
    
    For k = Col_Bound.Low To Col_Bound.High
        arr(i, k) = arr(j, k)
    Next k
End Function
Private Function ISort_AssignmentValue(i As Long, Value As Variant) As Variant
    Dim k As Long
    
    For k = Col_Bound.Low To Col_Bound.High
        arr(i, k) = Value(k)
    Next k
End Function
Private Function ISort_ReAssignmentValue(i As Long, Value As Variant) As Variant
    Dim k As Long
    
    ReDim Value(Col_Bound.Low To Col_Bound.High) As Variant
    For k = Col_Bound.Low To Col_Bound.High
        Value(k) = arr(i, k)
    Next k
End Function

'三数取中
Private Function ISort_MedianOfThree(ByVal Low As Long, ByVal High As Long) As Variant
    Dim m As Long
    Dim Value() As Variant '返回某一行的1维数组
    Dim k As Long
    
    m = Low + (High - Low) / 2
    
    If arr(Low, i_col) > arr(High, i_col) Then Call ISort_Swap(Low, High)  '交换左端与右端数据，保证左端较小
    If arr(m, i_col) > arr(High, i_col) Then Call ISort_Swap(High, m)      '交换中间与右端数据，保证中间较小
    If arr(m, i_col) > arr(Low, i_col) Then Call ISort_Swap(m, Low)        '交换中间与左端数据，保证左端为中间值
    
    ISort_ReAssignmentValue Low, Value
    
    ISort_MedianOfThree = Value
End Function
