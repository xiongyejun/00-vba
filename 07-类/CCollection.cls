VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'有n组二维数组，根据关键字的列ColKey，求这些数据的并集或交集

Private Type DataItem
    arr() As Variant
    
    LowCol As Long
    HighCol As Long
    Cols As Long
    Low As Long
    High As Long
End Type

Private Type Data
    Items() As DataItem
    iLen As Long
    iNext As Long
    Rows As Long    '所有数据的行数，这个是有可能有重复的在里面的
    
    ColKey As Long  '关键字列
    ValCol As Long 'UnionValues的时候，需要把哪列数据用来对比
    Cols As Long '每个数据的列，每一列的内容必须一致
End Type

Private d As Data

Private Const ErrNoValue As String = "没有添加任何数据。"
Private Const ErrNotSetColKey As String = "没有设置关键字列。"
Private Const ErrNotSetValCol As String = "没有设置对比的数据列。"
Private Const ErrColKeyOutBounds As String = "关键字列越界。"
Private Const ErrValColOutBounds As String = "对比数据列越界。"
Private Const ErrNotSameCols As String = "数据的列不一致。"

Property Let ColKey(Value As Long)
    d.ColKey = Value
End Property
Property Let ValCol(Value As Long)
    d.ValCol = Value
End Property


'添加数据
Sub AddData(Value() As Variant, RowStart As Long, RowEnd As Long)
    If d.iNext = d.iLen Then
        RedimItems
    End If
    
    d.Items(d.iNext).arr = Value
    d.Items(d.iNext).Low = RowStart
    d.Items(d.iNext).High = RowEnd
    d.Rows = RowEnd - RowStart + 1 + d.Rows
    
    d.Items(d.iNext).LowCol = LBound(Value, 2)
    d.Items(d.iNext).HighCol = UBound(Value, 2)
    d.Items(d.iNext).Cols = d.Items(d.iNext).HighCol - d.Items(d.iNext).LowCol + 1
    
    d.iNext = d.iNext + 1
End Sub

'求并集,并去重复的
Function Union(ReturnArr() As Variant, ByRef ReturnIndex As Long, Optional bUnionValues As Boolean = False) As String
    Dim strCheck As String
    '检查数据
    strCheck = Check()
    If VBA.Len(strCheck) Then
        Union = strCheck
        Exit Function
    End If
    
    If bUnionValues Then
        Union = UnionValues(ReturnArr, ReturnIndex)
    Else
        Union = OnlyUnion(ReturnArr, ReturnIndex)
    End If
End Function
'求并集,仅仅是把数据合并到一起，原来10列，合并后还是10列，只是合并同类项了
Private Function OnlyUnion(ReturnArr() As Variant, ByRef ReturnIndex As Long) As String
    'Excel处理的数据一般不会很大，直接定义个足够大的数组
    ReDim ReturnArr(d.Rows - 1, d.Items(0).LowCol To d.Items(0).HighCol) As Variant
    
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, iCol As Long, iRow As Long, p_row As Long
    Dim strKey As String
    ReturnIndex = 0
    For i = 0 To d.iNext - 1
        For iRow = d.Items(i).Low To d.Items(i).High
            strKey = VBA.CStr(d.Items(i).arr(iRow, d.ColKey))
            If Not dic.Exists(strKey) Then
                dic(strKey) = ReturnIndex
                '复制数据
                For iCol = d.Items(i).LowCol To d.Items(i).HighCol
                    ReturnArr(ReturnIndex, iCol) = d.Items(i).arr(iRow, iCol)
                Next iCol
                ReturnIndex = ReturnIndex + 1
            
            Else
                '存在就累加数据
                p_row = dic(strKey)
                ReturnArr(p_row, d.ValCol) = VBA.CDbl(d.Items(i).arr(iRow, d.ValCol)) + ReturnArr(p_row, d.ValCol)
            End If
        Next
    Next
    
    Set dic = Nothing
End Function
'求并集,合并的时候，同时把每组数据里的valCol分别放置
Private Function UnionValues(ReturnArr() As Variant, ByRef ReturnIndex As Long) As String
    Dim ReturnCols As Long, p_col As Long, p_row As Long
    ReturnCols = d.Cols - 1 + d.iNext
    
    ReDim ReturnArr(d.Rows - 1, ReturnCols) As Variant
    
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, iCol As Long, iRow As Long
    Dim strKey As String
    ReturnIndex = 0
    For i = 0 To d.iNext - 1
        For iRow = d.Items(i).Low To d.Items(i).High
            strKey = VBA.CStr(d.Items(i).arr(iRow, d.ColKey))
            If dic.Exists(strKey) Then
                '存在的时候，只需要加ValCol那列的数据
                '定位好行就行，添加ValCol那列的数据存在不存在都要做的
                p_row = dic(strKey)
            Else
                dic(strKey) = ReturnIndex
                p_row = ReturnIndex
                '复制数据
                p_col = 0
                For iCol = d.Items(i).LowCol To d.Items(i).HighCol
                    If iCol <> d.ValCol Then
                        ReturnArr(ReturnIndex, p_col) = d.Items(i).arr(iRow, iCol)
                        p_col = p_col + 1
                    End If
                    
                Next iCol
                ReturnIndex = ReturnIndex + 1
            End If
            
            p_col = d.Cols - 1 + i
            '累加数据
            ReturnArr(p_row, p_col) = VBA.CDbl(d.Items(i).arr(iRow, d.ValCol)) + ReturnArr(p_row, p_col)
        Next
    Next
    
    Set dic = Nothing
End Function
'求交集，并去重复的
Function Intersection() As String
    Dim strCheck As String
    '检查数据
    strCheck = Check()
    If VBA.Len(strCheck) Then
        Intersection = strCheck
        Exit Function
    End If
    
End Function

'检查数据
Private Function Check() As String
    Dim i As Long
    
    If d.iNext = 0 Then
        Check = ErrNoValue '没有数据
        Exit Function
    ElseIf d.ColKey = -1 Then
        Check = ErrNotSetColKey '没有设置关键字列
        Exit Function
    ElseIf d.ColKey < d.Items(0).LowCol Or d.ColKey > d.Items(0).HighCol Then '关键字列越界
        Check = ErrColKeyOutBounds
        Exit Function
    ElseIf d.ValCol = -1 Then
        Check = ErrNotSetValCol
        Exit Function
    ElseIf d.ValCol < d.Items(0).LowCol Or d.ValCol > d.Items(0).HighCol Then
        Check = ErrValColOutBounds
        Exit Function
    End If
    '对比每组数据的列是否相等
    For i = 0 To d.iNext - 2
        If d.Items(i).Cols <> d.Items(i + 1).Cols Then
            Check = ErrNotSameCols
            Exit Function
        End If
    Next i
    d.Cols = d.Items(0).Cols
    
    Check = ""
End Function


Private Function RedimItems()
    If d.iLen < 1024 Then
        d.iLen = d.iLen * 1.2
    Else
        d.iLen = d.iLen * 2
    End If
    
    ReDim Preserve d.Items(d.iLen - 1) As DataItem 'item type
End Function

Private Sub Class_Initialize()
    d.iLen = 10
    d.iNext = 0
    ReDim d.Items(d.iLen - 1) As DataItem 'item type
    
    d.ColKey = -1
    d.ValCol = -1
End Sub


'    Dim i_row As Long
'
'    ActiveSheet.AutoFilterMode = False
'    i_row = Cells(Cells.Rows.Count, 1).End(xlUp).Row
'    If i_row < 2 Then MsgBox "没有数据": Exit Sub
'
'    Dim c_coll As New CCollection
'    Dim arr() As Variant
'
'    arr = Range("A1").Resize(i_row, 3).Value
'    c_coll.AddData arr, 1, i_row
'
'    arr = Range("D1").Resize(i_row, 3).Value
'    c_coll.AddData arr, 1, i_row
'
'    arr = Range("g1").Resize(i_row, 3).Value
'    c_coll.AddData arr, 1, i_row
'
'    c_coll.ColKey = 1
'    c_coll.ValCol = 2
'
'    Dim result() As Variant, resultIndex As Long
'    Dim strReturn As String
'
'    strReturn = c_coll.Union(result, resultIndex, True)
'
'    If VBA.Len(strReturn) Then
'        MsgBox strReturn
'    Else
'        With Range("J1").Resize(resultIndex, UBound(result, 2) - LBound(result, 2) + 1)
'            .EntireColumn.Clear
'            .Value = result
'        End With
'    End If
