Attribute VB_Name = "CAnt"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'蚂蚁类

'1 初始化 , 蚂蚁从某个城市出发
'2 选择下一个城市
'3 移动到下一个城市
'4 如果全部城市都去过了，直接返回到出发城市，到第5步，否则再跳转到第2步
'5 计算走过的路径长度

'用列表来操作是否去过的城市
Private Type ListStruct
    Head As Long
    Arr() As Long '下标是城市，值是下一个城市
    pArr() As Double '概率
End Type

Private Type DataStruct
    CityCount As Long ' 城市数量
    P() As Double       '城市间的概率
    L As ListStruct
    
    AllowedCity() As Byte
End Type
Private d As DataStruct

Property Let CityCount(Value As Long)
    d.CityCount = Value
    
    ReDim d.AllowedCity(d.CityCount - 1) As Byte
End Property
'初始化还没去过的城市
Private Function InitAllowedCity() As Long
    Dim i As Long
    
    For i = 0 To d.CityCount - 1
        d.AllowedCity(i) = 1
    Next
End Function
'返回路径的长度
'P  城市间的概率，每次都要更新
Function Go(ByRef Path() As Long, ByRef P() As Double, ByRef Distance() As Double) As Double
    d.P = P
    InitAllowedCity
    
    '随机选择1个开始的城市
    Path(0) = SelectRndCity()
    d.AllowedCity(Path(0)) = 0
    DelListNode Path(0) - 1, Path(0)
    
    Dim i As Long, tmpLenth As Double
    For i = 1 To d.CityCount - 1 - 1
        Path(i) = GetNext(Path(i - 1), P)
        '记录去过的城市
        d.AllowedCity(Path(i)) = 0
        '计算走的路程
        tmpLenth = tmpLenth + Distance(Path(i - 1), Path(i))
    Next
    '最后一个是唯一的，不需要选择
    Path(i) = d.L.Head
    d.AllowedCity(Path(i)) = 0
    tmpLenth = tmpLenth + Distance(Path(i - 1), Path(i))
    
    '加上返回的路程
    tmpLenth = tmpLenth + Distance(Path(i), Path(0))
    Go = tmpLenth
End Function

Private Function GetNext(iFrom As Long, ByRef P() As Double) As Long
    Dim i As Long
    
    '计算总的信息素概率
    Dim pTotal As Double
    
    i = d.L.Head
    Do Until i = -1
        pTotal = pTotal + P(iFrom, i)
        i = d.L.Arr(i)
    Loop
    
    If pTotal > 0# Then
        '计算每个可以去的城市的概率
        i = d.L.Head
        Do Until i = -1
            d.L.pArr(i) = P(iFrom, i) / pTotal
            i = d.L.Arr(i)
        Loop
        
        '轮盘选择
        i = DiskSelect()
    Else
        i = d.L.Head
        DelListNode -1, i
    End If
    
    GetNext = i
End Function

Private Function SelectRndCity() As Long
    VBA.Randomize
    
    SelectRndCity = VBA.Rnd() * (d.CityCount - 1)
End Function
'轮盘选择
Private Function DiskSelect() As Long
    Dim i As Long, iPre As Long
    Dim P As Double
    
    VBA.Randomize
    P = VBA.Rnd()
    
    iPre = -1
    i = d.L.Head
    P = P - d.L.pArr(i)
    Do Until P < 0#
        iPre = i
        i = d.L.Arr(i)
        P = P - d.L.pArr(i)
    Loop
    '选中后从列表中删除
    DelListNode iPre, i
    
    DiskSelect = i
End Function
'初始化列表
Function InitList() As Long
    Dim i As Long
    
    ReDim d.L.Arr(d.CityCount - 1) As Long
    ReDim d.L.pArr(d.CityCount - 1) As Double
    For i = 0 To d.CityCount - 1 - 1
        d.L.Arr(i) = i + 1
    Next
    d.L.Arr(i) = -1 '结束标志-1
    d.L.Head = 0
End Function
'列表中删除1个节点
Private Function DelListNode(iPre As Long, iNow As Long) As Long
    If iPre = -1 Then
        '头节点删除
        d.L.Head = d.L.Arr(iNow)
    Else
        d.L.Arr(iPre) = d.L.Arr(iNow)
    End If
End Function


