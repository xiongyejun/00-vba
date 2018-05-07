Attribute VB_Name = "CTSP"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'TSP（traveling salesman problem旅行商问题）

Private Type DataStruct
    Alpha As Long '   启发因子，信息素的重要程度，一般取值1.0。
    Beta As Long '    期望因子，城市间距离的重要程度，一般取值2.0。
    Rou As Double ' 信息素挥发系数，一般取值0.5。
    AntCount As Long '  蚂蚁数量，一般取值为城市数量的2/3。
    CityCount As Long ' 城市数量
    ItCount  As Long ' 迭代次数，就是全部蚂蚁搜索多少次，取值自己设定。
    Q  As Long ' 信息素总量，取值多少对算法没什么影响。
    
    Distance() As Double '城市距离
    Info() As Double     '信息素
    P() As Double       '城市间的概率
End Type

'蚂蚁揍的路径和路径的总长度
Private Type PathInfo
    Arr() As Long  '记录的是第几个城市
    Lenth As Double '路径的长度
End Type
'记录蚂蚁的信息，n个蚂蚁，n条路径
Private Type AntInfo
    Ants() As CAnt
    Path() As PathInfo '蚂蚁走的路径
    
    Best As PathInfo
End Type

Private d As DataStruct
Private ant As AntInfo

Function Go() As Long
    Dim i As Long, j As Long
    Dim t As Double, tAnt As Double, tmpT As Double
    
    t = Timer
    
    ant.Best.Lenth = 999999999999#
    ReDim ant.Best.Arr(d.CityCount - 1) As Long
    '多次迭代
    For i = 0 To d.ItCount - 1
        If i Mod 10 = 0 Then
            Application.StatusBar = VBA.Format(i, "0/") & VBA.Format(d.ItCount, "0……")
            DoEvents
        End If
        '每只蚂蚁走动
        For j = 0 To d.AntCount - 1
            tmpT = Timer
            ant.Ants(j).InitList
            ant.Path(j).Lenth = ant.Ants(j).Go(ant.Path(j).Arr, d.P, d.Distance)
            '记录最好的路径
            If ant.Path(j).Lenth < ant.Best.Lenth Then
                ant.Best.Lenth = ant.Path(j).Lenth
                ant.Best.Arr = ant.Path(j).Arr
            End If
            tAnt = tAnt + Timer - tmpT
        Next
        
        '更新环境信息素
        UpdateInfo
        '信息素更新后，更新城市间概率
        CountProbability
    Next
    Application.StatusBar = False

    Debug.Print t, Timer, Timer - t, tAnt
    '输出最后的pbest
    PrintOut
End Function
'更新环境信息素
Private Function UpdateInfo() As Long
    Dim i As Long, j As Long
    '先挥发
    For i = 0 To d.CityCount - 1
        For j = i + 1 To d.CityCount - 1
            d.Info(i, j) = d.Info(i, j) * (1 - d.Rou)
            d.Info(j, i) = d.Info(i, j)
        Next
    Next
    
    Dim tmpInfo As Double
    '加每个蚂蚁留下的信息素
    For i = 0 To d.AntCount - 1
        '蚂蚁在路径上留下的信息素=Q/L
        tmpInfo = d.Q / ant.Path(i).Lenth
        For j = 0 To d.CityCount - 1 - 1
            d.Info(ant.Path(i).Arr(j), ant.Path(i).Arr(j + 1)) = d.Info(ant.Path(i).Arr(j), ant.Path(i).Arr(j + 1)) + tmpInfo
            d.Info(ant.Path(i).Arr(j + 1), ant.Path(i).Arr(j)) = d.Info(ant.Path(i).Arr(j), ant.Path(i).Arr(j + 1))
        Next
    Next
End Function

Private Function PrintOut() As Long
    Dim Arr() As String
    
    ReDim Arr(d.CityCount - 1) As String
    Dim i As Long
    For i = 0 To d.CityCount - 1
        Arr(i) = VBA.CStr(ant.Best.Arr(i))
    Next
    
    Debug.Print ant.Best.Lenth, VBA.Join(Arr, "→")
End Function

'迭代次数
Property Let ItCount(Value As Long)
    d.ItCount = Value
End Property
'城市数量
Private Function SetCityCount()
    d.CityCount = UBound(d.Distance) + 1
    d.AntCount = d.CityCount * 3 \ 2
    
    Dim i As Long, j As Long
    
    '初始蚂蚁
    ReDim ant.Ants(d.AntCount - 1) As CAnt
    ReDim ant.Path(d.AntCount - 1) As PathInfo
    For i = 0 To d.AntCount - 1
        Set ant.Ants(i) = New CAnt
        ant.Ants(i).CityCount = d.CityCount
        
        ReDim ant.Path(i).Arr(d.CityCount - 1) As Long
    Next
    
    '初始信息素
    ReDim d.Info(d.CityCount - 1, d.CityCount - 1) As Double
    For i = 0 To d.CityCount - 1
        For j = 0 To d.CityCount - 1
            d.Info(i, j) = 1#
        Next
    Next
    
    '初始概率
    ReDim d.P(d.CityCount - 1, d.CityCount - 1) As Double
    CountProbability
End Function
'城市间距离
Property Let Distance(Value() As Double)
    d.Distance = Value
    
    SetCityCount
End Property

'计算城市间的概率
Private Function CountProbability()
    Dim i As Long, j As Long
    
    For i = 0 To d.CityCount - 1
        For j = i + 1 To d.CityCount - 1
            d.P(i, j) = Probability(i, j)
            d.P(j, i) = d.P(i, j)
        Next
    Next
End Function

'计算到某个城市的概率
Private Function Probability(iFrom As Long, iTo As Long) As Double
    Probability = (d.Info(iFrom, iTo) ^ d.Alpha) / (d.Distance(iFrom, iTo) ^ d.Beta)
End Function

Private Sub Class_Initialize()
    d.Alpha = 1
    d.Beta = 5
    d.Rou = 0.5
    d.Q = 100
End Sub

Private Sub Class_Terminate()
    Erase d.Distance, d.P, d.Info
    Erase ant.Ants, ant.Path
End Sub
