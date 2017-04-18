VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLChildRsibling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Left Child Right Sibling
Const NULL_VAL As Long = &H80000000

Private Type TreeNode
    Value As Long
    Left As Long
    Right As Long
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

'NodeValue  要增加的节点的值  对应的就是在数组中的下标，也就是表格的行号
'AddIndex   要增加在那个节点的下面
Function Insert(NodeValue As Long, AddIndex As Long) As Long
    Dim tmp_index As Long
   
    tmp_index = ArrAlloc()
    If ArrNode(AddIndex).Left = NULL_VAL Then
        ArrNode(AddIndex).Left = tmp_index
    Else
        '找到右孩子是空的那个节点
        AddIndex = ArrNode(AddIndex).Left
        Do Until ArrNode(AddIndex).Right = NULL_VAL
            AddIndex = ArrNode(AddIndex).Right
        Loop
       
        ArrNode(AddIndex).Right = tmp_index
    End If
    ArrNode(tmp_index).Left = NULL_VAL
    ArrNode(tmp_index).Right = NULL_VAL
    ArrNode(tmp_index).Value = NodeValue
   
    Insert = tmp_index '返回具体是在哪个节点下增加了
End Function

Sub PrintOut(arr() As Long)
    Dim k As Long
    Dim i_level As Long
   
    ReDim arr(Me.MaxSize - 1, 1) As Long '0列是数据源的行号，1列是对应的层次
    k = 0
    i_level = 0
   
    PrintTree Root, k, arr, i_level
End Sub

Private Function PrintTree(NodeIndex As Long, ByRef k As Long, arr() As Long, ByVal i_level As Long)
    If NodeIndex = NULL_VAL Then Exit Function
   
    arr(k, 0) = ArrNode(NodeIndex).Value: arr(k, 1) = i_level
    k = k + 1
    PrintTree ArrNode(NodeIndex).Left, k, arr, i_level + 1
    PrintTree ArrNode(NodeIndex).Right, k, arr, i_level
End Function

'到数组中找1个位置放新增的节点
Private Function ArrAlloc() As Long
    ArrAlloc = cls_stack.Pop()
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
        ArrNode(i).Value = NULL_VAL
        ArrNode(i).Left = NULL_VAL
        ArrNode(i).Right = NULL_VAL
    Next i
   
    '根节点初始化，第1层次的图号都挂在不存在的root下
    '增加根节点
    Root = ArrAlloc()
    If Root = -1 Then Exit Function '栈空了，没有了可用的空间
    ArrNode(Root).Left = NULL_VAL
    ArrNode(Root).Right = NULL_VAL
    ArrNode(Root).Value = 1 '指向原表格的第1行，也就是标题行
End Function
