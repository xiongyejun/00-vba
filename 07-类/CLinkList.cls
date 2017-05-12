VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLinkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const NULL_VAL As Long = &H80000000

Private Enum CompareMode
    Small = -1
    LargeEqual = 1
End Enum

Private Type Node
    Data As Long
    Next As Long
End Type

Private Arr() As Node
Private ArrSize As Long     '数组的最大个数
Private Head As Long
Private NodeNum As Long     '当前Node的数量
Private cls_stack As CStack      '记录Node中空闲位置

Property Let MaxSize(Value As Long)
    ArrSize = Value
    ReDim Arr(ArrSize - 1) As Node
   
    InitArr
End Property

Property Get Count() As Long
    Count = NodeNum
End Property

'增加节点
Sub Insert(NodeValue As Long)
    Dim node_index As Long
    Dim node_pre As Long    '前1个的位置
    Dim node_next As Long    '后1个的位置
   
    If NodeNum = ArrSize Then
        MsgBox "链表已满。"
        Exit Sub
    End If
   
    node_index = ArrAlloc()
    Arr(node_index).Data = NodeValue
   
    node_next = Find(NodeValue, node_pre, LargeEqual)
    If node_next = NULL_VAL Then
        If node_pre = NULL_VAL Then
            '链表还是空
            Head = node_index
        Else
            Arr(node_pre).Next = node_index
        End If
        Arr(node_index).Next = NULL_VAL
    Else
        If node_pre = NULL_VAL Then
            Head = node_index
        Else
            Arr(node_pre).Next = node_index
        End If
        Arr(node_index).Next = node_next
    End If
   
    NodeNum = NodeNum + 1
End Sub

'删除节点
'删除成功返回节点的位置，否则返回NULL_VAL
Function Delete(NodeValue As Long) As Long
    Dim node_index As Long
    Dim node_pre As Long    '前1个的位置
   
    '目前这里是不对的
    node_index = Find(NodeValue, node_pre, LargeEqual)
   
    If node_index = NULL_VAL Then
        Delete = NULL_VAL
    Else
        If node_pre = NULL_VAL Then
            '说明要删除的是头结点
            Head = Arr(node_index).Next
        Else
            Arr(node_pre).Next = Arr(node_index).Next
        End If
       
        cls_stack.Push (node_index) '记录删除的节点位置
        NodeNum = NodeNum - 1
    End If
End Function

'查找node
Private Function Find(NodeValue As Long, ByRef node_pre As Long, i_compare As CompareMode) As Long
    Dim node_index As Long
   
    node_pre = NULL_VAL
    If Head = NULL_VAL Then
        Find = NULL_VAL
        Exit Function
    End If
   
    node_index = Head
    Do Until (i_compare = ValueCompare(Arr(node_index).Data, NodeValue)) Or Arr(node_index).Next = NULL_VAL
        node_pre = node_index
        node_index = Arr(node_index).Next
    Loop
   
    If (i_compare = ValueCompare(Arr(node_index).Data, NodeValue)) Then
        Find = node_index
    Else
        node_pre = node_index
        Find = NULL_VAL
    End If
   
End Function
'-1     va11<val2
'01      va11>=val2
Private Function ValueCompare(val1 As Long, val2 As Long) As Long
    If val1 < val2 Then
        ValueCompare = -1
    Else
        ValueCompare = 1
    End If
End Function

Sub PrintLink()
    Dim node_index As Long
    Dim i As Long
   
    i = 1
    node_index = Head
    Do Until node_index = NULL_VAL
        Cells(i, 3).Value = Arr(node_index).Data
        node_index = Arr(node_index).Next
        i = i + 1
    Loop
End Sub

'到数组中找1个位置放新增的节点
Private Function ArrAlloc() As Long
    ArrAlloc = cls_stack.Pop()
End Function

Private Sub Class_Initialize()
    Me.MaxSize = 10
End Sub

Private Sub Class_Terminate()
    Erase Arr
    Set cls_stack = Nothing
End Sub

'初始化
Private Function InitArr()
    Set cls_stack = New CStack
    Dim i As Long
   
    Head = NULL_VAL
    NodeNum = 0
    cls_stack.MaxSize = ArrSize
   
    For i = 0 To ArrSize - 1
        cls_stack.Push i    '开始栈是满的，记录了可用的位置
        Arr(i).Data = NULL_VAL
        Arr(i).Next = NULL_VAL
    Next i
End Function
