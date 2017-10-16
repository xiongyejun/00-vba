Attribute VB_Name = "MCombinCols"
Option Explicit

'数据源是多个1维数组，多个1维数组之间的组合，有多少列就取多少个数的组合
'这种组合的数据同1个数组的数据不能出现2次
Type MyType
    i_rows As Long
    arr_data() As Variant
End Type

Type CombinColsDataType
    Arr() As String     '一组数据源
    Count As Long       '数据的个数
End Type

Type ReptCountType
    Self As Long        '自身重复的次数=其后各列相乘
    Col As Long         '整列需要循环的次数=其前各列相乘
End Type

Type CombinColsType
    Data() As CombinColsDataType     '多组数据源
    ColNum As Long                   '数据源的有多少组
    
    Result() As CombinColsDataType   '结果
    ReptCount() As ReptCountType     '循环次数控制
    ResultNum As Long                '结果的个数--等于各列个数相乘
    pResult As Long                  '指向正在生成的结果
End Type

Sub testCombinCol()
    Dim cbType As CombinColsType
    Dim i As Long, j As Long
    
    cbType.ColNum = 3
    ReDim cbType.Data(cbType.ColNum - 1) As CombinColsDataType
    '初始化数据
    cbType.Data(0).Count = 7
    ReDim cbType.Data(0).Arr(cbType.Data(0).Count - 1) As String
    For i = 0 To cbType.Data(0).Count - 1
        cbType.Data(0).Arr(i) = VBA.CStr(i)
    Next
    
    cbType.Data(1).Count = 2
    ReDim cbType.Data(1).Arr(cbType.Data(1).Count - 1) As String
    For i = 0 To cbType.Data(1).Count - 1
        cbType.Data(1).Arr(i) = VBA.Chr(i + VBA.Asc("a"))
    Next
    
    cbType.Data(2).Count = 3
    ReDim cbType.Data(2).Arr(cbType.Data(2).Count - 1) As String
    For i = 0 To cbType.Data(2).Count - 1
        cbType.Data(2).Arr(i) = VBA.Chr(i + VBA.Asc("A"))
    Next
    '计算每列数据应该重复的次数
    ReDim cbType.ReptCount(cbType.ColNum - 1) As ReptCountType
    cbType.ResultNum = 1
    For i = 0 To cbType.ColNum - 1
        cbType.ReptCount(i).Self = 1
        cbType.ReptCount(i).Col = 1
        
        cbType.ResultNum = cbType.ResultNum * cbType.Data(i).Count
        
        '自身重复的次数=其后各列相乘
        For j = i + 1 To cbType.ColNum - 1
            cbType.ReptCount(i).Self = cbType.ReptCount(i).Self * cbType.Data(j).Count
        Next
        
        '整列需要循环的次数=其前各列相乘
        For j = 0 To i - 1
            cbType.ReptCount(i).Col = cbType.ReptCount(i).Col * cbType.Data(j).Count
        Next
    Next
    '初始化结果数组
    ReDim cbType.Result(cbType.ResultNum - 1) As CombinColsDataType
    For i = 0 To cbType.ResultNum - 1
        ReDim cbType.Result(i).Arr(cbType.ColNum - 1) As String
    Next
    
    CombinCols cbType
    For i = 0 To cbType.ResultNum - 1
        Debug.Print i, VBA.Join(cbType.Result(i).Arr, "、")
    Next
End Sub

Function CombinCols(cbType As CombinColsType)
    Dim i As Long, j As Long, k As Long, m As Long

    '根据结果的index定位到数据的行，直接赋值
    For i = 0 To cbType.ResultNum - 1
        For j = 0 To cbType.ColNum - 1
            k = i \ cbType.ReptCount(j).Self
            m = k Mod cbType.Data(j).Count
            
            cbType.Result(i).Arr(j) = cbType.Data(j).Arr(m)
        Next j
    Next i
End Function

