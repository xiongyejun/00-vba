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

Private Root As CNode
Private z As CNode '尾节点用于表示空树

'增加节点
Sub Add(Value As Long)
'    Set Root = InsertR(Root, Value)
    Insert Value
End Sub
'调整平衡
Sub Balance()
    Set Root = BalanceR(Root)
End Sub
'选择第k小的值
Function SelectValue(k As Long) As Long
    SelectValue = SelectR(Root, k)
End Function

'recursion递归插入
Private Function InsertR(h As CNode, Value As Long) As CNode
    If h Is z Then
        Set InsertR = NewNode(Value, z, z, 1)
        Exit Function
    End If
    
    If Value < h.Value Then
        Set h.Left = InsertR(h.Left, Value)
    Else
        Set h.Right = InsertR(h.Right, Value)
    End If
    
    h.N = h.N + 1
    Set InsertR = h
End Function
'非递归插入
Private Function Insert(Value As Long)
    If Root Is z Then
        Set Root = NewNode(Value, z, z, 1)
        Exit Function
    End If
    
    Dim p As CNode 'x的parent
    Dim x As CNode '当前节点，找到1个空节点插入
    
    Set x = Root
    Do While Not x Is z
        Set p = x
        x.N = x.N + 1
        If Value < x.Value Then
            Set x = x.Left
        Else
            Set x = x.Right
        End If
    Loop
    
    '插入到父节点p下面
    Set x = NewNode(Value, z, z, 1)
    If Value < p.Value Then
        Set p.Left = x
    Else
        Set p.Right = x
    End If
End Function
'右旋转rotation right
Private Function RotR(h As CNode) As CNode
    Dim x As CNode
    
    Set x = h.Left
    Set h.Left = x.Right
    Set x.Right = h
    
    '只有x和h的子树有变化，所有只需要更新x和h的N
    'h和她的右子树下移一层
    h.N = h.Left.N + h.Right.N + 1
    'x和它的左子树上移一层
    x.N = x.Left.N + x.Right.N + 1
    
    Set RotR = x
End Function
Private Function RotL(h As CNode) As CNode
    Dim x As CNode
    
    Set x = h.Right
    Set h.Right = x.Left
    Set x.Left = h
    
    h.N = h.Left.N + h.Right.N + 1
    x.N = x.Left.N + x.Right.N + 1
        
    Set RotL = x
End Function

'递归划分操作
Private Function PartR(h As CNode, k As Long) As CNode
    If h Is z Then
        Set PartR = z
        Exit Function
    End If
    
    Dim tmp As Long
    tmp = h.Left.N
    
    If tmp > k Then
        Set h.Left = PartR(h.Left, k)
        Set h = RotR(h)
    ElseIf tmp < k Then
        Set h.Right = PartR(h.Right, k - tmp - 1)
        Set h = RotL(h)
    End If
    Set PartR = h
End Function
'递归的调整平衡，将中值节点放在树根
Private Function BalanceR(h As CNode) As CNode
    If h.N < 2 Then
        Set BalanceR = h
        Exit Function
    End If
    
    '将中值节点放在树根
    Set h = PartR(h, h.N \ 2)
    '递归处理左右子树
    Set h.Left = BalanceR(h.Left)
    Set h.Right = BalanceR(h.Right)
    Set BalanceR = h
End Function

'选择第k小的值，下标是0，数据升序排序后的下标k
Private Function SelectR(h As CNode, k As Long) As Long
    If h Is z Then
        SelectR = NULL_VAL
        Exit Function
    End If
    
    If h.Left.N > k Then
        SelectR = SelectR(h.Left, k)
        Exit Function
    End If
    
    If h.Left.N < k Then
        SelectR = SelectR(h.Right, k - h.Left.N - 1)
        Exit Function
    End If
    
    SelectR = h.Value
End Function

Private Function NewNode(Value As Long, l As CNode, r As CNode, N As Long) As CNode
    Set NewNode = New CNode
    NewNode.Value = Value
    NewNode.N = N
    Set NewNode.Left = l
    Set NewNode.Right = r
End Function

Sub PrintOut()
    PrintTree Root
End Sub

Private Function PrintTree(N As CNode)
    If N Is Nothing Then Exit Function

    Debug.Print N.Value
    PrintTree N.Left
    PrintTree N.Right
End Function

Sub DrawTree()
    Dim p_root As Point

    p_root.x = 400
    p_root.y = 10
    DrawNode Root, p_root, 0
End Sub

Private Function DrawNode(N As CNode, p As Point, ByVal i_level As Long)
    Dim arr() As Long
    Dim tmp_p As Point

    If N Is Nothing Then Exit Function

    arr = getChildNode(p, 100, 155 - i_level * 15, 25 + i_level * 15)

    DrawOval p.x, p.y, N.Value
    If Not N.Left Is Nothing Then drawLine p.x + 10, p.y + 17, arr(0, 0) + 10, arr(0, 1), 1
    If Not N.Right Is Nothing Then drawLine p.x + 10, p.y + 17, arr(1, 0) + 10, arr(1, 1), 1

    tmp_p.x = arr(0, 0): tmp_p.y = arr(0, 1)
    DrawNode N.Left, tmp_p, i_level + 1
    tmp_p.x = arr(1, 0): tmp_p.y = arr(1, 1)
    DrawNode N.Right, tmp_p, i_level + 1
End Function

Private Function Free(h As CNode)
    If h Is z Then
        Exit Function
    End If
    
    Free h.Left
    Free h.Right
    
    Set h.Left = Nothing
    Set h.Right = Nothing
    Set h = Nothing
End Function

Private Sub Class_Initialize()
    Set z = New CNode
    z.Value = NULL_VAL
    
    Set Root = z
End Sub

Private Sub Class_Terminate()
    Set z = Nothing
    Free Root
End Sub
