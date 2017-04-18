VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAVLTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const NULL_VAL As Long = &H80000000

Private Type TreeNode
    Value As Long
    Left As Long
    Right As Long
    Height As Long
End Type

Private ArrNode() As TreeNode   '保存树节点的数组
Private cls_stack As CStack      '记录ArrNode中空闲位置
Private TreeMaxSize As Long
Private Root As Long

Property Let MaxSize(Value As Long)
    TreeMaxSize = Value
    ReDim ArrNode(TreeMaxSize - 1) As TreeNode
   
    InitArr
End Property

Property Get MaxSize() As Long
    MaxSize = TreeMaxSize
End Property
'增加节点
Sub Add(NodeValue As Long)
    Insert NodeValue, Root
End Sub
'NodeValue  要增加的节点的值
'AddIndex   要增加的节点下标        对外部调用不需要的，只是递归时需要
Private Function Insert(NodeValue As Long, ByRef AddIndex As Long)
    If AddIndex = NULL_VAL Then
        AddIndex = ArrAlloc() '指向他的child
        If AddIndex = -1 Then Exit Function '栈空了，没有了可用的空间
       
        ArrNode(AddIndex).Left = NULL_VAL
        ArrNode(AddIndex).Right = NULL_VAL
        ArrNode(AddIndex).Value = NodeValue
        Exit Function
    End If
   
    If NodeValue < ArrNode(AddIndex).Value Then
        Insert NodeValue, ArrNode(AddIndex).Left
        
        If GetHeight(ArrNode(AddIndex).Left) - GetHeight(ArrNode(AddIndex).Right) = 2 Then
            If NodeValue < ArrNode(ArrNode(AddIndex).Left).Value Then
                AddIndex = SingleRotateWithLeft(AddIndex)
            Else
                AddIndex = DoubleRotateWithLeft(AddIndex)
            End If
        End If
'        Insert NodeValue, ArrNode(AddIndex).Left
    ElseIf NodeValue > ArrNode(AddIndex).Value Then
        Insert NodeValue, ArrNode(AddIndex).Right
        
        If GetHeight(ArrNode(AddIndex).Right) - GetHeight(ArrNode(AddIndex).Left) = 2 Then
            If NodeValue > ArrNode(ArrNode(AddIndex).Right).Value Then
                AddIndex = SingleRotateWithRight(AddIndex)
            Else
                AddIndex = DoubleRotateWithRight(AddIndex)
            End If
        End If
'        Insert NodeValue, ArrNode(AddIndex).Right
    'Else   '相等就不进行操作
    End If
    
    ArrNode(AddIndex).Height = Application.WorksheetFunction.Max(GetHeight(ArrNode(AddIndex).Left), GetHeight(ArrNode(AddIndex).Right)) + 1
    
    Insert = AddIndex
End Function

Private Function SingleRotateWithLeft(k2 As Long) As Long
    Dim k1 As Long
    
    k1 = ArrNode(k2).Left
    ArrNode(k2).Left = ArrNode(k1).Right
    ArrNode(k1).Right = k2
    
    ArrNode(k2).Height = Application.WorksheetFunction.Max(GetHeight(ArrNode(k2).Left), GetHeight(ArrNode(k2).Right)) + 1
    ArrNode(k1).Height = Application.WorksheetFunction.Max(GetHeight(ArrNode(k1).Left), ArrNode(k2).Height) + 1
    
    SingleRotateWithLeft = k1
End Function

Private Function SingleRotateWithRight(k2 As Long) As Long
    Dim k1 As Long
    
    k1 = ArrNode(k2).Right
    ArrNode(k2).Right = ArrNode(k1).Left
    ArrNode(k1).Left = k2
    
    ArrNode(k2).Height = Application.WorksheetFunction.Max(GetHeight(ArrNode(k2).Right), GetHeight(ArrNode(k2).Left)) + 1
    ArrNode(k1).Height = Application.WorksheetFunction.Max(GetHeight(ArrNode(k1).Right), ArrNode(k2).Height) + 1
    
    SingleRotateWithRight = k1
End Function

Private Function DoubleRotateWithLeft(k3 As Long) As Long
    ArrNode(k3).Left = SingleRotateWithRight(ArrNode(k3).Left)
    DoubleRotateWithLeft = SingleRotateWithLeft(k3)
End Function

Private Function DoubleRotateWithRight(k3 As Long) As Long
    ArrNode(k3).Right = SingleRotateWithLeft(ArrNode(k3).Right)
    DoubleRotateWithRight = SingleRotateWithRight(k3)
End Function

Private Function GetHeight(NodeIndex) As Long
    If NodeIndex = NULL_VAL Then
        GetHeight = -1
    Else
        GetHeight = ArrNode(NodeIndex).Height
    End If
End Function


'删除节点
Function Remove(NodeValue As Long)
    Root = DeleteNode(Root, NodeValue)
End Function

Private Function DeleteNode(NodeIndex As Long, NodeValue As Long) As Long
    Dim tmp_index As Long
   
    If NodeIndex = NULL_VAL Then
        MsgBox "没找到要删除的元素。"
    ElseIf NodeValue < ArrNode(NodeIndex).Value Then
        ArrNode(NodeIndex).Left = DeleteNode(ArrNode(NodeIndex).Left, NodeValue)
    ElseIf NodeValue > ArrNode(NodeIndex).Value Then
        ArrNode(NodeIndex).Right = DeleteNode(ArrNode(NodeIndex).Right, NodeValue)
    '下面上找到了元素
    ElseIf ArrNode(NodeIndex).Left <> NULL_VAL And ArrNode(NodeIndex).Right <> NULL_VAL Then
        '含有左右child
        tmp_index = FindMin(ArrNode(NodeIndex).Right)
        ArrNode(NodeIndex).Value = ArrNode(tmp_index).Value
        ArrNode(tmp_index).Value = NodeValue
        ArrNode(NodeIndex).Right = DeleteNode(ArrNode(NodeIndex).Right, NodeValue)
    Else
        tmp_index = NodeIndex
        If ArrNode(NodeIndex).Left = NULL_VAL Then
            NodeIndex = ArrNode(NodeIndex).Right
        ElseIf ArrNode(NodeIndex).Right = NULL_VAL Then
            NodeIndex = ArrNode(NodeIndex).Left
        End If
        cls_stack.Push tmp_index 'free释放空间
    End If
   
    DeleteNode = NodeIndex
End Function
'找到最小值所在的位置
Function FindMin(NodeIndex As Long) As Long
    If NodeIndex = NULL_VAL Then
        FindMin = NULL_VAL
    ElseIf ArrNode(NodeIndex).Left = NULL_VAL Then
        FindMin = NodeIndex
    Else
        FindMin = FindMin(ArrNode(NodeIndex).Left)
    End If
End Function

'到数组中找1个位置放新增的节点
Private Function ArrAlloc() As Long
    ArrAlloc = cls_stack.Pop()
End Function

Sub PrintOut()
    PrintTree Root
End Sub

Private Function PrintTree(NodeIndex As Long)
    If NodeIndex = NULL_VAL Then Exit Function
   
    PrintTree ArrNode(NodeIndex).Left
    Debug.Print ArrNode(NodeIndex).Value
    PrintTree ArrNode(NodeIndex).Right
End Function

Sub DrawTree()
    Dim p_root As Point
   
    p_root.x = 300
    p_root.y = 10
    DrawNode Root, p_root, 0
End Sub

Private Function DrawNode(NodeIndex As Long, p As Point, ByVal i_level As Long)
    Dim arr() As Long
    Dim tmp_p As Point
   
    If NodeIndex = NULL_VAL Then Exit Function
   
    arr = getChildNode(p, 100, 155 - i_level * 15, 25 + i_level * 15)
   
    DrawOval p.x, p.y, ArrNode(NodeIndex).Value
    If ArrNode(NodeIndex).Left <> NULL_VAL Then drawLine p.x + 10, p.y + 17, arr(0, 0) + 10, arr(0, 1), 1
    If ArrNode(NodeIndex).Right <> NULL_VAL Then drawLine p.x + 10, p.y + 17, arr(1, 0) + 10, arr(1, 1), 1
   
    tmp_p.x = arr(0, 0): tmp_p.y = arr(0, 1)
    DrawNode ArrNode(NodeIndex).Left, tmp_p, i_level + 1
    tmp_p.x = arr(1, 0): tmp_p.y = arr(1, 1)
    DrawNode ArrNode(NodeIndex).Right, tmp_p, i_level + 1
End Function

Private Sub Class_Initialize()
    Me.MaxSize = 10
    Root = NULL_VAL
End Sub

Private Sub Class_Terminate()
    Erase ArrNode
   
    Set cls_stack = Nothing
End Sub
'初始化
Private Function InitArr()
    Set cls_stack = New CStack
    Dim i As Long
   
    cls_stack.MaxSize = Me.MaxSize
   
    For i = 0 To Me.MaxSize - 1
        cls_stack.Push i    '开始栈是满的，记录了可用的位置
        
        ArrNode(i).Left = NULL_VAL
        ArrNode(i).Right = NULL_VAL
        ArrNode(i).Value = NULL_VAL
    Next i
End Function


