VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHeap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'大顶堆数据结构——父节点大于等于2个子节点
'始终保持最大值在root上
'数据存在在数组中，以二叉树结构操作
'位置i的父节点位置在i\2处，子节点在2i和2i+1处

Private Const MAX_VAL As Long = &H7FFFFFFF

Private Type Node
    Key As Long
    Data As Long
End Type
'数组的最大个数
Private Max_Size As Long
'存放数据，数组中0下标的不放具体数据，用作观察哨
Private ArrNode() As Node
'指向ArrNode的下一个可用单元格
Private P_Arr As Long
'

Property Let MaxSize(Value As Long)
    Max_Size = Value
    
    DeapInit
End Property
'插入数据
Function Insert(Key As Long) As Long
    P_Arr = P_Arr + 1
    ArrNode(P_Arr).Key = Key
    FixUp
End Function
'删除最大值，返回最大值
Function DelMax() As Long
    Swap 1, P_Arr

    DelMax = ArrNode(P_Arr).Key
    '数组减少1个
    P_Arr = P_Arr - 1
    '重新调整堆，以满足堆的性质
    FixDown
End Function
'晋升--在P_Arr+1处插入1个新节点，然后调整节点与其父的位置，保证满足堆的性质
Private Function FixUp() As Long
    Dim i As Long '父节点
    Dim j As Long '子节点
    
    j = P_Arr
    i = j \ 2
    Do While Compare(i, j) = -1
        Swap j, i
        j = i
        i = i \ 2
    Loop
    
    FixUp = j
End Function
'删除最大值root后，P_Arr处的节点移动到root处，然后调整节点与其子节点的位置，保证满足堆的性质
Private Function FixDown(Optional index As Long = 1) As Long
    Dim i As Long '父节点
    Dim j As Long '子节点
    
    i = index
    Do While 2 * i <= P_Arr
        j = 2 * i '左子孩子
        '找到左右孩子中，大的那个
        If j < P_Arr And Compare(j, j + 1) = -1 Then j = j + 1
        If Compare(i, j) > -1 Then Exit Do
        Swap i, j
        i = j
    Loop
    
    FixDown = j
End Function
'交换2个节点的位置
Private Function Swap(Index1 As Long, Index2 As Long)
    Dim tmp As Node
    
    tmp = ArrNode(Index1)
    ArrNode(Index1) = ArrNode(Index2)
    ArrNode(Index2) = tmp
End Function
'对比2个节点的大小
'-1     va11<val2
'0      va11=val2
'1      va11>val2
Private Function Compare(Index1 As Long, Index2 As Long) As Long
    If ArrNode(Index1).Key < ArrNode(Index2).Key Then
        Compare = -1
    ElseIf ArrNode(Index1).Key = ArrNode(Index2).Key Then
        Compare = 0
    Else
        Compare = 1
    End If
End Function
'置空
Private Function EmptyDeap()
    P_Arr = 0
End Function
'堆的初始化
Private Function DeapInit()
    ReDim ArrNode(Max_Size) As Node
    
    ArrNode(0).Key = MAX_VAL
    P_Arr = 0
End Function
Sub PrintOut(arr() As Long)
    Dim i As Long
    
    For i = 1 To Max_Size
        arr(i, 1) = ArrNode(i).Key
    Next
End Sub

'Sub HeapSort(ArrSort() As Node, l As Long, r As Long)
'    Dim k As Long
'    Dim N As Long
'    Dim i As Long
'
'    N = r - l + 1
'    k = N \ 2 '跳过叶子节点
'    For i = k To 1 Step -1
'        FixDown i
'    Next
'
'    Do While N > 1
'        Swap 1, N
'        FixDown 1
'        N = N - 1
'    Loop
'End Sub
