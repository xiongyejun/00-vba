VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'����QQ 648555205
'�Ա����� https://item.taobao.com/item.htm?spm=686.1000925.0.0.H2w0Qj&id=534896704968

Private arr() As Long
Private StackMaxSize As Long
Private StackTop As Long        'ָ��Stack��1���հ�λ��

Property Get Top() As Long
    Top = StackTop
End Property
Property Let MaxSize(Value As Long)
    StackMaxSize = Value
    ReDim arr(StackMaxSize - 1) As Long
End Property
Property Get MaxSize() As Long
    MaxSize = StackMaxSize
End Property
'����
Function Pop() As Long
    If StackTop = 0 Then
        MsgBox "��ջ��"
        Pop = -1
        Exit Function
    End If
   
    StackTop = StackTop - 1
    Pop = arr(StackTop)
End Function
'ѹ��
Function Push(Value As Long)
    If StackTop = Me.MaxSize Then
        MsgBox "��ջ��"
        Exit Function
    End If
   
    arr(StackTop) = Value
    StackTop = StackTop + 1
End Function
'���
Function EmptyStack()
    StackTop = 0
End Function

Private Sub Class_Initialize()
    Me.MaxSize = 10
    StackTop = 0
End Sub
Private Sub Class_Terminate()
    Erase arr
End Sub
