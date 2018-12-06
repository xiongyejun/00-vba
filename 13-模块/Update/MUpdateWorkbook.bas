Attribute VB_Name = "MUpdateWorkbook"
Option Explicit

Private Enum Pos
    RowStart = 2
    
    SrcWk = 2
    SrcSht
    SrcRowStart
    SrcColKey
    SrcRng
    
    DesWk
    DesSht
    DesRowStart
    DesColKey
    DesRng
    
    TheType
    BeiZhu
End Enum

Private Enum UpdateType
    rng       '单纯的单元格对单元格赋值
    ColRelation    '按列的对应关系来一列一列的赋值
    ColRelationAppend '追加--和上面的不同在于赋值前不要进行清除
    Formula     '直接写个公式进去
    dic         '按colKey记录到dic，写入,没有找到的就为空
    dicExists    '同上，但如果没有找到的，就保留原来的内容
    AddName     '添加自定义名称---目标工作簿、工作表、目标单元格记录的都是文件名称，都需要添加
                'RowStart和ColKey是long类型，不记录
    
    AddNo '添加序号
End Enum
'excel文件之间的更新
Private Type DataStructItem
    wkName As String
    shtName As String
    Action As String        '记录操作的str
                            'Rng           记录的是单元格地址
                            'ColRelation   记录的是列的对应关系-ColRelation
                            'Formula       记录的是公式，目标单元格记录单元格，源单元格写公式
                            'dic           记录的是item的列
                            
    RowStart As Long        '数据行开始的位置
    ColKey As Long          '定位用的列
End Type

'记录工作簿，源工作簿不需要保存，目标工作簿需要保存
Private Type WkType
    wk As Workbook
    bSave As Boolean
    wkName As String
End Type

Private Type DataStruct
    wk As Workbook      '记录对应关系的wk
    sht As Worksheet    '记录对应关系的sht
    Path As String
    Rows As Long
    Arr() As Variant
    
    Count As Long
    Src() As DataStructItem
    Des() As DataStructItem
    uType() As UpdateType

    dicWk As Object '字典记录工作表的名称--对应ArrWk的下标
    ArrWk() As WkType
End Type

'RangeByCol记录列对应关系的分隔符
Private Const SPLIT_WORD As String = "、"

Sub UpdateWorkbook()
    Dim d As DataStruct
    
    On Error GoTo err_handle
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set d.wk = ActiveWorkbook
    Set d.sht = ActiveSheet
    d.Path = d.wk.Path & "\"
    
    If ReturnCode.ErrRT = ReadData(d) Then Exit Sub
    If ReturnCode.ErrRT = DataToStruct(d) Then Exit Sub
    
    If ReturnCode.ErrRT = OpenAllWk(d) Then Exit Sub
    
    If ReturnCode.ErrRT = GetResult(d) Then
        If VBA.MsgBox("是否关闭所有工作簿？", vbYesNo) = vbNo Then Exit Sub
    End If
    
    If ReturnCode.ErrRT = CloseAllWk(d) Then Exit Sub
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "OK"
    
    Exit Sub
err_handle:
    MsgBox Err.Description
    If VBA.MsgBox("是否关闭所有工作簿？", vbYesNo) = vbYes Then CloseAllWk d
End Sub

Private Function GetResult(d As DataStruct) As ReturnCode
    Dim i As Long
    
    For i = 0 To d.Count - 1
'        If i = 2 Then Stop
        
        Select Case d.uType(i)
        Case UpdateType.rng
            If ReturnCode.ErrRT = GetResultRange(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
            
        Case UpdateType.ColRelation
            If ReturnCode.ErrRT = GetResultColRelation(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
            
        Case UpdateType.ColRelationAppend
            If ReturnCode.ErrRT = GetResultColRelationAppend(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
            
        Case UpdateType.Formula
            If ReturnCode.ErrRT = GetResultFormula(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
        
        Case UpdateType.dic
            If ReturnCode.ErrRT = GetResultDic(d, i, False) Then
                GetResult = ErrRT
                Exit Function
            End If
            
        Case UpdateType.dicExists
            If ReturnCode.ErrRT = GetResultDic(d, i, True) Then
                GetResult = ErrRT
                Exit Function
            End If
        
        Case UpdateType.AddName
            If ReturnCode.ErrRT = GetResultAddName(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
        
        Case UpdateType.AddNo
            If ReturnCode.ErrRT = GetResultAddNo(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
             
        Case Else
            MsgBox "未知类型" & VBA.CStr(d.uType(i))
            GetResult = ErrRT
            Exit Function
        End Select
    Next
    
    GetResult = SuccessRT
End Function

'添加序号
Private Function GetResultAddNo(d As DataStruct, i As Long) As ReturnCode
    Dim shtDes As Worksheet
    Dim RowDes As Long
    
    '目标工作表
    Set shtDes = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName)
    shtDes.AutoFilterMode = False
    RowDes = shtDes.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row
    If RowDes < d.Des(i).RowStart Then
        d.sht.Cells(i, Pos.BeiZhu + 1).Value = "Des数据没有超过RowStart"
        '仅提示，不返回错误
        GetResultAddNo = SuccessRT
        Exit Function
    End If
    
    With shtDes.Cells(d.Des(i).RowStart, d.Des(i).Action)
        .Value = 1
        .AutoFill Destination:=.Resize(RowDes - d.Des(i).RowStart + 1, 1), Type:=xlFillSeries
    End With
    
End Function


'添加自定义名称
Private Function GetResultAddName(d As DataStruct, i As Long) As ReturnCode
    '读取数据源
    Dim shtSrc As Worksheet, RowSrc As Long
    Dim ArrKey(), ArrValue()
    
    Set shtSrc = d.ArrWk(d.dicWk(d.Src(i).wkName)).wk.Worksheets(d.Src(i).shtName)
    shtSrc.AutoFilterMode = False
    RowSrc = shtSrc.Cells(Cells.Rows.Count, d.Src(i).ColKey).End(xlUp).Row
    If RowSrc < d.Src(i).RowStart Then
        d.sht.Cells(i, Pos.BeiZhu + 1).Value = "Src数据没有超过RowStart"
        '仅提示，不返回错误
        GetResultAddName = SuccessRT
        Exit Function
    Else
        ArrKey = shtSrc.Cells(1, d.Src(i).ColKey).Resize(RowSrc, 1).Value
        ArrValue = shtSrc.Cells(1, d.Src(i).Action).Resize(RowSrc, 1).Value
         
        If VBA.Len(d.Des(i).wkName) Then AddNameToWk ArrKey, ArrValue, d.ArrWk(d.dicWk(d.Des(i).wkName)).wk
        If VBA.Len(d.Des(i).shtName) Then AddNameToWk ArrKey, ArrValue, d.ArrWk(d.dicWk(d.Des(i).shtName)).wk
        If VBA.Len(d.Des(i).Action) Then AddNameToWk ArrKey, ArrValue, d.ArrWk(d.dicWk(d.Des(i).Action)).wk
    End If

    GetResultAddName = SuccessRT
End Function

Private Function AddNameToWk(ArrKey(), ArrValue(), wk As Workbook)
    Dim i As Long
    Dim strKey As String
    Dim strValue As String
    
    For i = 2 To UBound(ArrKey)
        strKey = VBA.CStr(ArrKey(i, 1))
        strValue = VBA.CStr(ArrValue(i, 1))
        If VBA.Len(strValue) = 0 Then strValue = "0"
        
        If VBA.Len(strKey) Then
            wk.Names.Add Name:=strKey, RefersToR1C1:="=" & strValue
        End If
    Next
End Function

Private Function GetResultDic(d As DataStruct, i As Long, bDicExists As Boolean) As ReturnCode
    '源表定位列是key，单元格列是item
    Dim shtDes As Worksheet
    Dim shtSrc As Worksheet
    Dim dic As Object, j As Long, strKey As String
    Dim ArrKey(), ArrItem()
    
    Set dic = CreateObject("Scripting.Dictionary")
    
    Dim RowDes As Long, RowSrc As Long
    '找到数据源的范围
    Set shtSrc = d.ArrWk(d.dicWk(d.Src(i).wkName)).wk.Worksheets(d.Src(i).shtName)
    shtSrc.AutoFilterMode = False
    RowSrc = shtSrc.Cells(Cells.Rows.Count, d.Src(i).ColKey).End(xlUp).Row
    If RowSrc < d.Src(i).RowStart Then
        d.sht.Cells(i, Pos.BeiZhu + 1).Value = "Src数据没有超过RowStart"
        '仅提示，不返回错误
        GetResultDic = SuccessRT
        Exit Function
    Else
        '目标工作表
        Set shtDes = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName)
        shtDes.AutoFilterMode = False
        RowDes = shtDes.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row
        If RowDes < d.Des(i).RowStart Then
            d.sht.Cells(i, Pos.BeiZhu + 1).Value = "Des数据没有超过RowStart"
            '仅提示，不返回错误
            GetResultDic = SuccessRT
            Exit Function
        End If
    
        ArrKey = shtSrc.Cells(1, d.Src(i).ColKey).Resize(RowSrc, 1).Value
        ArrItem = shtSrc.Cells(1, d.Src(i).Action).Resize(RowSrc, 1).Value
        For j = d.Src(i).RowStart To RowSrc
            dic(VBA.UCase$(VBA.CStr(ArrKey(j, 1)))) = ArrItem(j, 1)
        Next
        
        '输出
        ArrKey = shtDes.Cells(1, d.Des(i).ColKey).Resize(RowDes, 1).Value
        ArrItem = shtDes.Cells(1, d.Des(i).Action).Resize(RowDes, 1).Value
        For j = d.Des(i).RowStart To RowDes
            strKey = VBA.CStr(ArrKey(j, 1))
            strKey = VBA.UCase$(strKey)
            If dic.Exists(strKey) Then
                ArrItem(j, 1) = dic(strKey)
            Else
                If Not bDicExists Then
                    '没有就清空原来的内容
                    ArrItem(j, 1) = ""
                End If
            End If
        Next
        
        shtDes.Cells(1, d.Des(i).Action).Resize(RowDes, 1).Value = ArrItem
    End If
    
    GetResultDic = SuccessRT
    
    Set dic = Nothing
    Set shtDes = Nothing
    Set shtSrc = Nothing
End Function

Private Function GetResultColRelation(d As DataStruct, i As Long) As ReturnCode
    '这个需要先清除一下Des
    Dim shtDes As Worksheet
    Dim RowDes As Long
    
    Set shtDes = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName)
    shtDes.AutoFilterMode = False
    RowDes = shtDes.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row
    If RowDes >= d.Des(i).RowStart Then
        shtDes.Rows(VBA.CStr(d.Des(i).RowStart) & ":" & VBA.CStr(RowDes)).ClearContents
    End If
    
    GetResultColRelation = GetResultColRelationAppend(d, i)
End Function
Private Function GetResultColRelationAppend(d As DataStruct, i As Long) As ReturnCode
    Dim tmpSrc, tmpDes
    Dim shtDes As Worksheet
    Dim shtSrc As Worksheet
    
    tmpSrc = VBA.Split(d.Src(i).Action, SPLIT_WORD)
    tmpDes = VBA.Split(d.Des(i).Action, SPLIT_WORD)
    
    Dim iCount As Long
    iCount = UBound(tmpSrc) + 1
    If iCount <> UBound(tmpDes) + 1 Then
        MsgBox "列没有一一对应。" & vbNewLine & "出错所在行：" & VBA.CStr(i)
        GetResultColRelationAppend = ErrRT
        Exit Function
    End If
    
    Dim j As Long
    Dim RowDes As Long, RowSrc As Long
    '找到数据源的范围
    Set shtSrc = d.ArrWk(d.dicWk(d.Src(i).wkName)).wk.Worksheets(d.Src(i).shtName)
    shtSrc.AutoFilterMode = False
    RowSrc = shtSrc.Cells(Cells.Rows.Count, d.Src(i).ColKey).End(xlUp).Row
    If RowSrc < d.Src(i).RowStart Then
        d.sht.Cells(i, Pos.BeiZhu + 1).Value = "Src数据没有超过RowStart"
        '仅提示，不返回错误
        GetResultColRelationAppend = SuccessRT
        Exit Function
    Else
        '找到目标工作表的输出起始行
        Set shtDes = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName)
        shtDes.AutoFilterMode = False
        RowDes = shtDes.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row + 1
        If RowDes < d.Des(i).RowStart Then
            RowDes = d.Des(i).RowStart
        End If
        
        For j = 0 To iCount - 1
            shtDes.Cells(RowDes, VBA.CStr(tmpDes(j))).Resize(RowSrc - d.Src(i).RowStart + 1, 1).Value = shtSrc.Cells(d.Src(i).RowStart, VBA.CStr(tmpSrc(j))).Resize(RowSrc - d.Src(i).RowStart + 1, 1).Value
        Next
    End If
    
    '设置一下格式
'    shtDes.Range(shtDes.Cells(d.Des(i).RowStart, 1), shtDes.Cells(1, 1).CurrentRegion.SpecialCells(xlCellTypeLastCell)).Borders.LineStyle = 1


    GetResultColRelationAppend = SuccessRT
End Function

Private Function GetResultFormula(d As DataStruct, i As Long) As ReturnCode
    '如果d.Des(i).Action记录的仅是列号，则根据RowStart和ColKey来定位
    '判断右边第1个是否是数字
    Dim sht As Worksheet
    Dim rng As Range
    Dim i_row As Long
    
    Set sht = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName)
    If VBA.IsNumeric(VBA.Right$(d.Des(i).Action, 1)) Then
        Set rng = sht.Range(d.Des(i).Action)
    Else
        sht.AutoFilterMode = False
        i_row = sht.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row
        If i_row < d.Des(i).RowStart Then
            d.sht.Cells(i, Pos.BeiZhu + 1).Value = "数据没有超过RowStart"
            '仅提示，不返回错误
            GetResultFormula = SuccessRT
            Exit Function
        Else
            Set rng = sht.Cells(d.Des(i).RowStart, d.Des(i).Action).Resize(i_row - d.Des(i).RowStart + 1, 1)
        End If
    End If
    
    If d.Src(i).Action = "=SUM" Then
        sht.AutoFilterMode = False
        i_row = sht.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row
        
        Set rng = sht.Range(d.Des(i).Action)
        '统计rng本列到最后的单元格
        rng.Formula = "=SUM(" & Cells(d.Des(i).RowStart, rng.Column).Resize(i_row - d.Des(i).RowStart + 1, 1).Address & ")"
    Else
        rng.FormulaR1C1Local = d.Src(i).Action
    End If
    
    GetResultFormula = SuccessRT
End Function

Private Function GetResultRange(d As DataStruct, i As Long) As ReturnCode
    '检查下目标单元格范围与src单元格范围是否一致
    '不一致的情况下提示一下
    Dim RngSrc As Range, RngDes As Range
    Dim iRows1 As Long, iCols1 As Long
    Dim iRows2 As Long, iCols2 As Long
    
    On Error GoTo ErrHandle
    Set RngSrc = d.ArrWk(d.dicWk(d.Src(i).wkName)).wk.Worksheets(d.Src(i).shtName).Range(d.Src(i).Action)
    Set RngDes = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName).Range(d.Des(i).Action)
    
    iRows1 = RngSrc.Rows.Count
    iCols1 = RngSrc.Columns.Count
    
    iRows2 = RngDes.Rows.Count
    iCols2 = RngDes.Columns.Count
    
    If iRows1 <> iRows2 Or iCols1 <> iCols2 Then
        d.sht.Cells(i + Pos.RowStart, Pos.BeiZhu + 1).Value = "单元格范围不一致"
        Set RngDes = RngDes.Range("A1").Resize(iRows1, iCols1)
    End If
        
    RngDes.Value = RngSrc.Value
    
    GetResultRange = SuccessRT
    Exit Function
    
    '有可能单元格地址写错了
ErrHandle:
    MsgBox Err.Description & vbNewLine & "出错所在行：" & VBA.CStr(i + Pos.RowStart)
    d.wk.Activate
    Cells(i, 1).Select
    
    GetResultRange = ErrRT
End Function

Private Function ReadData(d As DataStruct) As ReturnCode
    ActiveSheet.AutoFilterMode = False
    d.Rows = Cells(Cells.Rows.Count, Pos.DesWk).End(xlUp).Row
    If d.Rows < Pos.RowStart Then
        MsgBox "没有数据"
        ReadData = ReturnCode.ErrRT
        Exit Function
    End If
    d.Arr = Cells(1, 1).Resize(d.Rows, Pos.BeiZhu).Value
    '清空下备注后面的一列，这一列会用来记录一些提示信息，比如2个单元格范围不一致
    d.sht.Cells(1, Pos.BeiZhu + 1).EntireColumn.Clear
    ReadData = ReturnCode.SuccessRT
End Function
'将数据放到结构体中去
Private Function DataToStruct(d As DataStruct) As ReturnCode
    Dim i As Long
    
    Dim dic As Object

    Set d.dicWk = CreateObject("Scripting.Dictionary")

    d.Count = d.Rows - Pos.RowStart + 1
    ReDim d.Src(d.Count - 1) As DataStructItem
    ReDim d.Des(d.Count - 1) As DataStructItem
    ReDim d.uType(d.Count - 1) As UpdateType
    Dim iTmp As Long
    
    For i = Pos.RowStart To d.Rows
        iTmp = i - Pos.RowStart
    
        d.Src(iTmp).wkName = VBA.CStr(d.Arr(i, Pos.SrcWk))
        d.Src(iTmp).shtName = VBA.CStr(d.Arr(i, Pos.SrcSht))
        d.Src(iTmp).RowStart = VBA.CLng(d.Arr(i, Pos.SrcRowStart))
        d.Src(iTmp).ColKey = VBA.CLng(d.Arr(i, Pos.SrcColKey))
        d.Src(iTmp).Action = VBA.CStr(d.Arr(i, Pos.SrcRng))
        
        d.Des(iTmp).wkName = VBA.CStr(d.Arr(i, Pos.DesWk))
        d.Des(iTmp).shtName = VBA.CStr(d.Arr(i, Pos.DesSht))
        d.Des(iTmp).RowStart = VBA.CLng(d.Arr(i, Pos.DesRowStart))
        d.Des(iTmp).ColKey = VBA.CLng(d.Arr(i, Pos.DesColKey))
        d.Des(iTmp).Action = VBA.CStr(d.Arr(i, Pos.DesRng))
        
        Select Case d.Arr(i, Pos.TheType)
        Case "Rng"
            d.uType(iTmp) = UpdateType.rng
        Case "ColRelation"
            d.uType(iTmp) = UpdateType.ColRelation
        Case "ColRelationAppend"
            d.uType(iTmp) = UpdateType.ColRelationAppend
        Case "Formula"
            d.uType(iTmp) = UpdateType.Formula
            If VBA.Left$(d.Src(iTmp).Action, 1) <> "=" Then d.Src(iTmp).Action = "=" & d.Src(iTmp).Action
        
        Case "dic"
            d.uType(iTmp) = UpdateType.dic
            
        Case "dicExists"
            d.uType(iTmp) = UpdateType.dicExists
        
        Case "AddName"
            d.uType(iTmp) = UpdateType.AddName
            '目标工作簿、工作表……记录的都是文件名称，都需要添加
            RecordWk d.dicWk, d.Des(iTmp).shtName
            RecordWk d.dicWk, d.Des(iTmp).Action
            
        Case "AddNo"
            d.uType(iTmp) = UpdateType.AddNo
            
        Case Else
            MsgBox "未知类型[" & Cells(i, Pos.TheType).Address(True, True) & "]"
            DataToStruct = ErrRT
            Exit Function
        End Select
        
        '记录所有工作簿名称
        RecordWk d.dicWk, d.Src(iTmp).wkName
        
        If VBA.Len(d.Des(iTmp).wkName) = 0 Or VBA.Len(d.Des(iTmp).shtName) = 0 Or VBA.Len(d.Des(iTmp).Action) = 0 Then
            DataToStruct = ErrRT
            MsgBox "目标工作簿、工作表、单元格不能都为空。" & "出错所在行：" & VBA.CStr(i)
            Exit Function
        End If
        
        RecordWk d.dicWk, d.Des(iTmp).wkName
    Next
      
    DataToStruct = SuccessRT
End Function
'记录所有工作簿名称
Private Function RecordWk(dic As Object, wkName As String)
    If VBA.Len(wkName) Then
        If Not dic.Exists(wkName) Then dic(wkName) = dic.Count
    End If
End Function


Private Function OpenAllWk(d As DataStruct) As ReturnCode
    '源工作簿和目标工作簿可能存在交叉情况
    Dim i As Long
    Dim strKey As String
    ReDim d.ArrWk(d.dicWk.Count - 1) As WkType
       
    For i = Pos.RowStart To d.Rows
        strKey = VBA.CStr(d.Arr(i, Pos.SrcWk))
        'SrcWk  有可能是空的，在formula情况下
        If VBA.Len(strKey) Then
            OpenWkItem d, strKey, False
        End If
        
        OpenWkItem d, VBA.CStr(d.Arr(i, Pos.DesWk)), True
        '在AddName状态下，目标工作表  目标单元格 都可以填写1个文件
        If d.Arr(i, Pos.TheType) = "AddName" Then
            OpenWkItem d, VBA.CStr(d.Arr(i, Pos.DesSht)), True
            OpenWkItem d, VBA.CStr(d.Arr(i, Pos.DesRng)), True
        End If
        
    Next i
    
    '打开目标工作簿，这里一定要设置bSave
    For i = Pos.RowStart To d.Rows
        strKey = VBA.CStr(d.Arr(i, Pos.SrcWk))
        
        d.ArrWk(OpenWkItem(d, VBA.CStr(d.Arr(i, Pos.DesWk)), True)).bSave = True
        '在AddName状态下，目标工作表  目标单元格 都可以填写1个文件
        If d.Arr(i, Pos.TheType) = "AddName" Then
            d.ArrWk(OpenWkItem(d, VBA.CStr(d.Arr(i, Pos.DesSht)), True)).bSave = True
            d.ArrWk(OpenWkItem(d, VBA.CStr(d.Arr(i, Pos.DesRng)), True)).bSave = True
        End If
        
    Next i
    
    OpenAllWk = SuccessRT
End Function

Private Function OpenWkItem(d As DataStruct, strKey As String, bSave As Boolean) As Long
    Dim pArr As Long
    
    pArr = d.dicWk(strKey)
    
    If d.ArrWk(pArr).wk Is Nothing Then
        If VBA.InStr(strKey, ":\") Then
            '不是当前文件夹下的文件
            Set d.ArrWk(pArr).wk = Workbooks.Open(strKey, False)
        Else
            '还没有打开
            Set d.ArrWk(pArr).wk = Workbooks.Open(d.Path & strKey, False)
            '目标工作簿最后需要保存
        End If
        d.ArrWk(pArr).wkName = strKey
        d.ArrWk(pArr).bSave = bSave
    End If
    
    OpenWkItem = pArr
End Function

Private Function CloseAllWk(d As DataStruct) As ReturnCode
    Dim i As Long
    
    For i = 0 To UBound(d.ArrWk)
        If Not d.ArrWk(i).wk Is Nothing Then
            d.ArrWk(i).wk.Close d.ArrWk(i).bSave
        End If
    Next
    
    CloseAllWk = SuccessRT
End Function
