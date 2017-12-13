VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'二维数组，多列排序，类似工作表排序

'对第x列进行排序后，要形成按x排序后，每个x类别的开始和结束行
Private Type LowHigh
    low As Long
    High As Long
End Type
'

Function Sort(Values() As Variant, sortCols() As Long) As Long
    Dim Low_High As LowHigh
    Low_High.low = LBound(Values, 1)
    Low_High.High = UBound(Values, 1)
    dgSort Values, Low_High, sortCols, LBound(sortCols)
    
End Function
'Low_High       前一列的范围
'indexSortCol   当前要排序的列的下标
Private Function dgSort(l() As Variant, Low_High As LowHigh, sortCols() As Long, ByVal indexSortCol As Long)
    Dim i_col As Long
    
    i_col = sortCols(indexSortCol)
    '先排序
    Dim tmp As LowHigh
    tmp.low = Low_High.low
    tmp.High = Low_High.High
    QuickSort l, i_col, tmp.low, tmp.High
    
    '是否已经排完了
    If indexSortCol >= UBound(sortCols) Then Exit Function
    
    '然后根据排序的这一列生成LowHigh数组，为后面的列排序
    Dim arrLowHigh() As LowHigh
    ReDim arrLowHigh(Low_High.High - Low_High.low) As LowHigh '完全不重复也就这么多了
    
    Dim p_row As Long
    Dim i As Long
    
    arrLowHigh(p_row).low = Low_High.low
    For i = Low_High.low + 1 To Low_High.High
        If l(i, i_col) <> l(i - 1, i_col) Then
            arrLowHigh(p_row).High = i - 1
            p_row = p_row + 1
            arrLowHigh(p_row).low = i
        End If
        
    Next i
    arrLowHigh(p_row).High = i - 1
    
    '完全不重复，后面没必要继续排序了
    If p_row = Low_High.High - Low_High.low Then Exit Function
    
    '递归对每组LowHigh进行后面的排序
    For i = 0 To p_row
        dgSort l, arrLowHigh(i), sortCols, indexSortCol + 1
    Next i
    
End Function

Function QuickSort(l() As Variant, iCol As Long, low As Long, High As Long)   '快速排序
    Dim Pivot As Long
    
    Do While low < High
        Pivot = MyPartition(l, iCol, low, High)
        QuickSort l, iCol, low, Pivot - 1    '对低子表递归排序
        low = Pivot + 1                       '尾递归
    Loop
End Function

Private Function MyPartition(l(), iCol As Long, ByVal low As Long, ByVal High As Long) As Long
    Dim PivotKey()                '枢轴
    
    MedianOfThree l, iCol, low, High, PivotKey '三数取中
    
    Do While low < High
        Do While low < High And l(High, iCol) >= PivotKey(iCol)
            High = High - 1
        Loop
'        l(Low) = l(High)
        ReplaceValue l, low, High     '采用替换而不是交换的方式进行操作
        
        Do While low < High And l(low, iCol) <= PivotKey(iCol)
            low = low + 1
        Loop
'        l(High) = l(Low)
        ReplaceValue l, High, low
    Loop
    
'    l(Low) = PivotKey
    Dim i As Long
    For i = LBound(l, 2) To UBound(l, 2)
        l(low, i) = PivotKey(i)
    Next
    
    MyPartition = low
End Function
'采用替换而不是交换的方式进行操作
Private Function ReplaceValue(l(), low As Long, High As Long)
    Dim i As Long
    
    For i = LBound(l, 2) To UBound(l, 2)
        l(low, i) = l(High, i)
    Next
End Function

Private Function Swap(l(), low As Long, High As Long)
    Dim iTemp
    Dim i As Long
    
    For i = LBound(l, 2) To UBound(l, 2)
        iTemp = l(low, i)
        l(low, i) = l(High, i)
        l(High, i) = iTemp
    Next
End Function

Private Function MedianOfThree(l(), iCol As Long, ByVal low As Long, ByVal High As Long, PivotKey())
    Dim m As Long
    
    m = low + (High - low) / 2
    
    If l(low, iCol) > l(High, iCol) Then Swap l, low, High '交换左端与右端数据，保证左端较小
    If l(m, iCol) > l(High, iCol) Then Swap l, High, m     '交换中间与右端数据，保证中间较小
    If l(m, iCol) > l(low, iCol) Then Swap l, m, low       '交换中间与左端数据，保证左端为中间值
    
    Dim i As Long
    ReDim PivotKey(LBound(l, 2) To UBound(l, 2))
    For i = LBound(l, 2) To UBound(l, 2)
        PivotKey(i) = l(low, i)
    Next
End Function
