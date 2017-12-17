VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'一维数组实现接口

Implements ISort

Private arr() As Variant

Property Let Data(Value As Variant)
    arr = Value
End Property
Property Get Data() As Variant
    Data = arr
End Property

'比较
Private Function ISort_Less(i As Long, j As Long) As Boolean
    ISort_Less = arr(i) < arr(j)
End Function
Private Function ISort_LessValue(i As Long, Value As Variant) As Boolean
    ISort_LessValue = arr(i) < Value
End Function
Private Function ISort_LagerValue(i As Long, Value As Variant) As Boolean
    ISort_LagerValue = arr(i) > Value
End Function

'交换
Private Function ISort_Swap(i As Long, j As Long) As Variant
    Dim tmp As Variant
    
    tmp = arr(i)
    arr(i) = arr(j)
    arr(j) = tmp
End Function

'赋值
Private Function ISort_Assignment(i As Long, j As Long) As Variant
    arr(i) = arr(j)
End Function
Private Function ISort_AssignmentValue(i As Long, Value As Variant) As Variant
    arr(i) = Value
End Function
Private Function ISort_ReAssignmentValue(i As Long, Value As Variant) As Variant
    Value = arr(i)
End Function

'三数取中
Private Function ISort_MedianOfThree(ByVal Low As Long, ByVal High As Long) As Variant
    Dim m As Long
    
    m = Low + (High - Low) / 2
    
    If arr(Low) > arr(High) Then Call ISort_Swap(Low, High)  '交换左端与右端数据，保证左端较小
    If arr(m) > arr(High) Then Call ISort_Swap(High, m)      '交换中间与右端数据，保证中间较小
    If arr(m) > arr(Low) Then Call ISort_Swap(m, Low)        '交换中间与左端数据，保证左端为中间值
    ISort_MedianOfThree = arr(Low)
End Function
