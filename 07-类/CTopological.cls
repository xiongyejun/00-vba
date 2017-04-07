VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTopological"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'在1个表示工程的有向图中，用顶点表示活动
'用弧表示活动之间的优先关系
'这样的有向图为顶点表示活动的网
'称为AOV网——Activity On Vertex Network

'在1个表示工程的带权有向图中，用顶点表示事件，用有向边表示活动
'用边上的权值表示活动的持续时间
'这种有向图的边表示活动的网
'称为AOE网——Activity On Edge Network


'边表节点
Private Type EdgeNode
    Adjvex As Long      '临接点域，存储该顶点对应的下标
    Weight As Long
    Next As Long
End Type

Private Type VertexNode
    In As Long            '顶点入度数量
    Data As String        '顶点域，存储顶点信息
    FirstEdge As Long     '边表头指针
End Type

Const NULL_VAL As Long = &H80000000
Dim arr_adjlist() As VertexNode
Dim arr_edgeNode() As EdgeNode
Dim VertexsNum As Long     '顶点的数量
Dim EdgesNum As Long       '边的数量
Dim ArrVertexInfo() As Long
Dim cls_stack As CStack  '记录入度为0的顶点

Dim etv() As Long, ltv() As Long  'earliest time of vertex事件最早发生时间和latest time of vertex最迟发生时间数组
Dim cls_stack_topo As CStack    '用于存储拓扑序列的栈

Property Let VertexInfo(arr() As Long)
    'VertexFrom    VertexTo    Weight
    'Arr需要按照VertexFrom进行排序，相同的放在一起
    ArrVertexInfo = arr
   
    GetVertexsEdgesNum
    InitVertexEdge
End Property
'根据ArrVertexInfo获取顶点和边的数量
Private Function GetVertexsEdgesNum()
    Dim i As Long
   
    '假定ArrVertexInfo的VertexFrom和VertexTo就是顶点的编号，所以只要找到最大值
    EdgesNum = 0: VertexsNum = 0
    For i = 0 To UBound(ArrVertexInfo, 1)
        If ArrVertexInfo(i, 1) >= 0 Then
            EdgesNum = EdgesNum + 1
        End If
        If VertexsNum < ArrVertexInfo(i, 0) Then VertexsNum = ArrVertexInfo(i, 0)
        If VertexsNum < ArrVertexInfo(i, 1) Then VertexsNum = ArrVertexInfo(i, 1)
    Next i
    VertexsNum = VertexsNum + 1
   
    ReDim arr_edgeNode(EdgesNum - 1) As EdgeNode
    ReDim arr_adjlist(VertexsNum - 1) As VertexNode
    '初始化顶点，如果1个顶点没有出度，可能不会在ArrVertexInfo中反应
    '不用担心最大顶点没有出现在ArrVertexInfo中，因为至少会有1个入度
    For i = 0 To VertexsNum - 1
        arr_adjlist(i).FirstEdge = NULL_VAL
        arr_adjlist(i).In = 0
        arr_adjlist(i).Data = "V" & VBA.CStr(i)
    Next i
   
    cls_stack.MaxSize = VertexsNum
    cls_stack_topo.MaxSize = VertexsNum
    ReDim etv(VertexsNum - 1) As Long
    ReDim ltv(VertexsNum - 1) As Long
End Function
'初始化顶点和边
Private Function InitVertexEdge()
    Dim i As Long
    Dim arr_edgeNode_index As Long
    Dim p_arr_edgeNode As Long
    Dim start_Vertex As Long
    Dim first_flag As Boolean
   
    p_arr_edgeNode = 0: first_flag = True: i = 0
    start_Vertex = ArrVertexInfo(0, 0)
    Do Until i > UBound(ArrVertexInfo, 1)
        'VertexTo可能为空的情况
        If ArrVertexInfo(i, 1) >= 0 Then
            If start_Vertex = ArrVertexInfo(i, 0) Then
                If first_flag Then
                    arr_edgeNode_index = p_arr_edgeNode: p_arr_edgeNode = p_arr_edgeNode + 1
                    arr_adjlist(start_Vertex).FirstEdge = arr_edgeNode_index
                    first_flag = False
                Else
                    arr_edgeNode(arr_edgeNode_index).Next = p_arr_edgeNode
                    arr_edgeNode_index = p_arr_edgeNode: p_arr_edgeNode = p_arr_edgeNode + 1
                End If
                'VertexTo的入度+1
                arr_adjlist(ArrVertexInfo(i, 1)).In = arr_adjlist(ArrVertexInfo(i, 1)).In + 1
                arr_edgeNode(arr_edgeNode_index).Adjvex = ArrVertexInfo(i, 1)
                arr_edgeNode(arr_edgeNode_index).Weight = ArrVertexInfo(i, 2)
                i = i + 1
            Else
                '前面1个的arr_edgeNode的Next放置NULL_VAL
                arr_edgeNode(arr_edgeNode_index).Next = NULL_VAL
                start_Vertex = ArrVertexInfo(i, 0)
                first_flag = True
            End If
        Else
            i = i + 1
        End If
    Loop
    arr_edgeNode(arr_edgeNode_index).Next = NULL_VAL
End Function
'拓扑排序
Function GetTopological() As String
    Dim i As Long
    Dim Vertex_index As Long, edge_index As Long
    Dim str As String
    Dim Vertex_out_count As Long    '统计输出的顶点数量
    Dim tmp_Vertex_index As Long
   
    '记录入度为0的顶点
    For i = 0 To VertexsNum - 1
        If arr_adjlist(i).In = 0 Then cls_stack.Push (i)
    Next i
   
    Vertex_out_count = 0
   
    Do Until cls_stack.Top = 0
        Vertex_index = cls_stack.Pop
        cls_stack_topo.Push (Vertex_index) '将弹出的顶点序号压入拓扑序列的栈
        str = str & arr_adjlist(Vertex_index).Data & "→"
        Vertex_out_count = Vertex_out_count + 1
       
        edge_index = arr_adjlist(Vertex_index).FirstEdge
        Do Until edge_index = NULL_VAL
            tmp_Vertex_index = arr_edgeNode(edge_index).Adjvex
            '它的前面顶点已经输出了，所以入度-1
            arr_adjlist(tmp_Vertex_index).In = arr_adjlist(tmp_Vertex_index).In - 1
            If arr_adjlist(tmp_Vertex_index).In = 0 Then cls_stack.Push (tmp_Vertex_index)
           
            If etv(Vertex_index) + arr_edgeNode(edge_index).Weight > etv(tmp_Vertex_index) Then etv(tmp_Vertex_index) = etv(Vertex_index) + arr_edgeNode(edge_index).Weight
           
            edge_index = arr_edgeNode(edge_index).Next
        Loop
    Loop
   
    If Vertex_out_count < VertexsNum Then
        GetTopological = "存在环，不能输出全部顶点。" & str
    Else
        GetTopological = str
    End If
End Function
'从源点到汇点具有最大长度的路径叫关键路径，在关键路径上的活动叫关键活动
Function GetCriticalPath() As String
    Dim i As Long
    Dim ete As Long, lte As Long '最早发生时间和最晚发生时间
    Dim Vertex_index As Long, edge_index As Long
    Dim tmp_Vertex_index As Long
    Dim str As String
   
    For i = 0 To VertexsNum - 1
        ltv(i) = etv(VertexsNum - 1)
    Next i
   
    Do Until cls_stack_topo.Top = 0
        Vertex_index = cls_stack_topo.Pop
       
        edge_index = arr_adjlist(Vertex_index).FirstEdge
        Do Until edge_index = NULL_VAL
            tmp_Vertex_index = arr_edgeNode(edge_index).Adjvex
            '求各顶点时间最晚发生时间
            If ltv(tmp_Vertex_index) - arr_edgeNode(edge_index).Weight < ltv(Vertex_index) Then ltv(Vertex_index) = ltv(tmp_Vertex_index) - arr_edgeNode(edge_index).Weight
            edge_index = arr_edgeNode(edge_index).Next
        Loop
    Loop
   
    '求ete、lte和关键活动
    For i = 0 To VertexsNum - 1
        edge_index = arr_adjlist(i).FirstEdge
        ete = etv(i)
       
        Do Until edge_index = NULL_VAL
            tmp_Vertex_index = arr_edgeNode(edge_index).Adjvex
            lte = ltv(tmp_Vertex_index) - arr_edgeNode(edge_index).Weight
            If ete = lte Then
                str = str & " <" & arr_adjlist(i).Data & "," & arr_adjlist(tmp_Vertex_index).Data & ">Weight=" & arr_edgeNode(edge_index).Weight
            End If
            edge_index = arr_edgeNode(edge_index).Next
        Loop
    Next i
    GetCriticalPath = str
End Function
Function PrintVertexInfo()
    Dim i As Long
    Dim tmp As Long
    Dim rng As Range
    Dim str As String
   
    Set rng = Range("E1")
    For i = 0 To VertexsNum - 1
        tmp = arr_adjlist(i).FirstEdge
        str = "In=" & arr_adjlist(i).In & " " & arr_adjlist(i).Data
       
        Do Until tmp = NULL_VAL
            str = str & "→" & arr_edgeNode(tmp).Adjvex
            tmp = arr_edgeNode(tmp).Next
        Loop
       
        rng.Offset(i, 0).Value = str
    Next i
End Function
Private Sub Class_Initialize()
    Set cls_stack = New CStack
    Set cls_stack_topo = New CStack
End Sub
Private Sub Class_Terminate()
    Set cls_stack = Nothing
    Set cls_stack = Nothing
    Erase arr_adjlist, arr_edgeNode
    Erase etv, ltv
End Sub
