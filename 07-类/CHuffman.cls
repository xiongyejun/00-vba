VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHuffman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'http://www.cnblogs.com/kubixuesheng/p/4397798.html
'1.根据给定的n个权值{w1,w2,…,wn}构成二叉树集合F={T1,T2,…,Tn},其中每棵二叉树Ti中只有一个带权为wi的根结点,其左右子树为空.
'2.在F中选取两棵根结点权值最小的树作为左右子树构造一棵新的二叉树,且置新的二叉树的根结点的权值为左右子树根结点的权值之和.
'3.在F中删除这两棵树,同时将新的二叉树加入F中.
'4.重复2、3,直到F只含有一棵树为止.(得到哈夫曼树)


'PKZip规定了最特殊的树，左边小，右边大
'可以直接倒推出来

Const NULL_VAL As Long = &H80000000

Private Type Node
'    Weight As Long  '权重
    Data As Long
    Left As Long
    Right As Long
    Parent As Long
    Height As Long
End Type

Private arr_data() As Long, arr_code_len() As Long
Private arr_node() As Node
Private p_arr_node As Long

Function SetArrData(ArrData() As Long, ArrCodeLen() As Long)
    Dim i As Long
    Dim k As Long
    Dim i_mid As Long
    
    arr_data = ArrData
    arr_code_len = ArrCodeLen
    k = UBound(arr_data) + 1
    '按照CodeLen码长进行降序排列
    InsertSort arr_code_len, arr_data, 0, k - 1
    '再按arr_val排序
    i = 0
    Do Until i >= k - 1
        i_mid = i + 1
        Do Until arr_code_len(i) <> arr_code_len(i_mid)
            i_mid = i_mid + 1
            If i_mid >= k Then Exit Do
        Loop
        
        If i_mid - i >= 1 Then InsertSort arr_data, arr_code_len, i, i_mid - 1
        
        i = i_mid
    Loop
    
    ReDim arr_node(k * 2) As Node
    p_arr_node = -1
    
    InitHuffman
End Function

Private Function InitHuffman()
    Dim i_code_len As Long
    Dim i_value As Long
    Dim i As Long
    Dim str_bit As String
    Dim tmp_node As Long    '记录要在哪个节点增加叶子
    Dim p_ArrCodeLen As Long
    
    p_ArrCodeLen = 0
    '初始右边那一条树
    For i = 1 To arr_code_len(0)
        AddNode NULL_VAL, NULL_VAL, i, p_arr_node, i
    Next i
    arr_node(p_arr_node).Right = NULL_VAL
    
    tmp_node = p_arr_node
    Do Until p_ArrCodeLen = UBound(arr_code_len) + 1
        '找到有空位置的节点
        Do Until arr_node(tmp_node).Left = NULL_VAL Or arr_node(tmp_node).Right = NULL_VAL
            tmp_node = arr_node(tmp_node).Parent
        Loop
        '如果这个节点的高度不够，就在它下面插入1个节点
        Do Until arr_node(tmp_node).Height = arr_code_len(p_ArrCodeLen)
            AddNode NULL_VAL, NULL_VAL, NULL_VAL, tmp_node, arr_node(tmp_node).Height + 1
            If arr_node(tmp_node).Right = NULL_VAL Then
                arr_node(tmp_node).Right = p_arr_node
            Else
                arr_node(tmp_node).Left = p_arr_node
            End If
            
            tmp_node = p_arr_node
        Loop
        '找到了可以增加叶子结点的地方了
        AddNode arr_data(p_ArrCodeLen), NULL_VAL, NULL_VAL, tmp_node, arr_node(tmp_node).Height + 1
        If arr_node(tmp_node).Right = NULL_VAL Then
            arr_node(tmp_node).Right = p_arr_node
        Else
            arr_node(tmp_node).Left = p_arr_node
        End If
        'ArrCodeLen已经没有作用，记录下他所在的节点index
        arr_code_len(p_ArrCodeLen) = p_arr_node
        p_ArrCodeLen = p_ArrCodeLen + 1
    Loop
    
End Function

'将bit代表的数字保存在字典中
Sub PrintBit(dic As Object)
    Dim i As Long
    Dim k As Long
    Dim str_bit As String
    Dim node_index As Long
    Dim left_index As Long
    
    k = UBound(arr_data) + 1
    
    For i = 0 To k - 1
        str_bit = ""
        
        left_index = arr_code_len(i)
        node_index = arr_node(arr_code_len(i)).Parent
        Do Until node_index = -1
            If arr_node(node_index).Left = left_index Then
                str_bit = "0" & str_bit
            Else
                str_bit = "1" & str_bit
            End If
            '记录当前的节点下标，方便判断是左还是右
            left_index = node_index
            node_index = arr_node(node_index).Parent
        Loop
        
'        Debug.Print arr_data(i), str_bit
        dic(str_bit) = arr_data(i)
    Next i
    
End Sub

Private Function AddNode(ByVal NodeValue As Long, ByVal LeftIndex As Long, ByVal RightIndex As Long, ByVal ParentIndex As Long, i_height As Long)
    p_arr_node = p_arr_node + 1
    
    arr_node(p_arr_node).Data = NodeValue
    arr_node(p_arr_node).Left = LeftIndex
    arr_node(p_arr_node).Right = RightIndex
    arr_node(p_arr_node).Parent = ParentIndex
    arr_node(p_arr_node).Height = i_height
End Function

Private Function InsertSort(arr_code_len() As Long, arr_val() As Long, Low As Long, High As Long)
    Dim i As Long, j As Long
    Dim ShaoBing As Long, ShaoBing_tmp As Long
    
    '先按arr_code_len排序
    For i = Low + 1 To High
        If arr_code_len(i) > arr_code_len(i - 1) Then
            ShaoBing = arr_code_len(i)             '设置哨兵
            ShaoBing_tmp = arr_val(i)
            
            j = i - 1
            Do While arr_code_len(j) < ShaoBing
                arr_code_len(j + 1) = arr_code_len(j)
                arr_val(j + 1) = arr_val(j)
                j = j - 1
                If j = Low - 1 Then Exit Do
            Loop
            
            arr_code_len(j + 1) = ShaoBing
            arr_val(j + 1) = ShaoBing_tmp
        End If
    Next i

End Function


'Private Function AddNode(NodeValue As Long, WeightValue As Long, LeftIndex As Long, RightIndex As Long)
'    p_arr_node = p_arr_node + 1
'    arr_node(p_arr_node).Weight = WeightValue
'    arr_node(p_arr_node).Data = NodeValue
'    arr_node(p_arr_node).Left = LeftIndex
'    arr_node(p_arr_node).Right = RightIndex
'    arr_node(p_arr_node).Parent = NULL_VAL
'End Function

Private Sub Class_Initialize()
    
End Sub

Private Sub Class_Terminate()
    Erase arr_data
End Sub
