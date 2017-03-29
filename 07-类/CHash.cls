VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const MUL_FACTOR As Long = 211

Private Type HashNode
    Key As String
    Item As String
End Type

Private arr() As HashNode
Private ArrSize As Long     '数组的最大个数
Private KeyCount As Long    '实际的key的个数
Private KeyMaxCount As Long '最大的key的个数

Const dSCALE As Double = 1.2 '数组放大的比例

Property Let MaxCount(Value As Long)
    KeyMaxCount = Value
    ArrSize = Value * dSCALE
   
    Do Until IsPrime(ArrSize)
        ArrSize = ArrSize + 1
    Loop
   
    ReDim arr(ArrSize) As HashNode
End Property

Property Get Count() As Long
    Count = KeyCount
End Property

Sub Add(Key As String, Optional Item As String = "")
    Dim i_index As Long
   
    i_index = GetIndex(Key)
   
    '找到hash对应的空位置（添加），'或者是找到已经存在的key（替换Item）
    If VBA.Len(arr(i_index).Key) = 0 Then
        If KeyCount >= KeyMaxCount Then
            MsgBox "已达到Key的最大数量，不能添加。"
            Exit Sub
        End If
        KeyCount = KeyCount + 1
       
        arr(i_index).Key = Key
    End If
   
    arr(i_index).Item = Item
End Sub

Function GetItem(Key As String) As String
    Dim i_index As Long
   
    i_index = GetIndex(Key)
    '没找到的时候添加Key
    If VBA.Len(arr(i_index).Key) = 0 Then Add (Key)
    GetItem = arr(i_index).Item
End Function

Function RemoveItem(Key As String)
    Dim i_index As Long
   
    i_index = GetIndex(Key)
    If VBA.Len(arr(i_index).Key) Then KeyCount = KeyCount - 1
    arr(i_index).Key = ""
    arr(i_index).Item = ""
End Function

Function Keys()

End Function

Function Exists(Key As String) As Boolean
    Dim i_index As Long
   
    i_index = GetIndex(Key)
    Exists = (Key = arr(i_index).Key)
End Function
'找到Key所在数组Arr的下标，或者是找到hash对应的空位置
Private Function GetIndex(Key As String) As Long
    Dim i_index As Long
   
    i_index = Hash(Key) Mod ArrSize
   
    Do Until VBA.Len(arr(i_index).Key) = 0 Or arr(i_index).Key = Key
        i_index = Collision(i_index)
    Loop
    GetIndex = i_index
End Function
'处理冲突的方法
Private Function Collision(i_index As Long) As Long
    Collision = (i_index + 1) Mod ArrSize
End Function

Private Function Hash(Key As String) As Long
    Dim i As Long
    Const m As Long = 7158271

    Hash = 5381
    For i = 1 To Len(Key)
        Hash = (MUL_FACTOR * Hash + AscW(Mid$(Key, i, 1))) Mod m
    Next i
    Hash = Math.Abs(Hash)
End Function

'除留余数法
'f(key) = key mod p (p<=m) m是散列表长
'p通常为小于或等于m的最小质数
'Function Hash(key_num As Long, p As Long) As Long
'    Hash = key_num * MUL_FACTOR Mod p
'End Function

Private Function IsPrime(num As Long) As Boolean '判断一个数是否是质数
    Dim temp As Long
    Dim i As Long
   
    If num Mod 2 Then
        temp = Math.Sqr(num)
        For i = 2 To temp
            If (num Mod i) = 0 Then
                IsPrime = False
                Exit Function
            End If
        Next i
        IsPrime = True
    Else
        IsPrime = False
    End If
End Function

Private Sub Class_Initialize()
    Me.MaxCount = 10
End Sub

Private Sub Class_Terminate()
    Erase arr
End Sub

