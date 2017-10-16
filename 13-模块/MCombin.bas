Attribute VB_Name = "MCombin"
Option Explicit

'从1个DataNum个的一维数组中，挑选ChooseNum个的组合结果

Type CombinResultItem
    Arr() As String     '1个结果就是1个chooseNum个数的数组
End Type

Type CombinType
    Data() As String    '一组数据源
    DataNum As Long     '数据源的个数
    ChooseNum As Long   '要选择多少个的组合
    
    Result() As CombinResultItem  '结果
    ResultNum As Long   '结果的个数，可以用工作表函数Combin计算
    pResult As Long     '指向正在生成的结果
End Type

Sub test()
    Dim cbType As CombinType
    Dim i As Long
    Const NUM_DATA As Long = 10
    
    '初始化数据
    ReDim cbType.Data(NUM_DATA - 1) As String
    For i = 0 To NUM_DATA - 1
        cbType.Data(i) = VBA.CStr(i)
    Next i
    
    cbType.DataNum = NUM_DATA
    cbType.ChooseNum = 2
    cbType.ResultNum = Application.WorksheetFunction.Combin(NUM_DATA, cbType.ChooseNum)
    
    '初始化结果数组
    ReDim cbType.Result(cbType.ResultNum - 1) As CombinResultItem
    For i = 0 To cbType.ResultNum - 1
        ReDim cbType.Result(i).Arr(cbType.ChooseNum - 1) As String
    Next
    '开始组合
    DGCombin cbType, 0, 0
    '打印组合结果
    For i = 0 To cbType.ResultNum - 1
        Debug.Print i, VBA.Join(cbType.Result(i).Arr, "、")
    Next
End Sub
'pData        '指向正要使用的数据源下标
'pChooseNum   '组合到第几个了
Function DGCombin(cbType As CombinType, pData As Long, pChooseNum As Long)
    Dim i As Long
    
    If pChooseNum = cbType.ChooseNum Then
        cbType.pResult = cbType.pResult + 1
        Exit Function
    End If
    
    cbType.Result(cbType.pResult).Arr(pChooseNum) = cbType.Data(pData)
    DGCombin cbType, pData + 1, pChooseNum + 1
    
    '剩下数据的个数，大于剩下还需要的数据个数，可以再生成组合
    If cbType.DataNum - pData > cbType.ChooseNum - pChooseNum Then
        'pChooseNum 之前的数据先要复制过来
        For i = 0 To pChooseNum - 1
            cbType.Result(cbType.pResult).Arr(i) = cbType.Result(cbType.pResult - 1).Arr(i)
        Next
        DGCombin cbType, pData + 1, pChooseNum
    End If
End Function

