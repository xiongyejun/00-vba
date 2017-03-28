VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Head As Long
Private Tail As Long

Private QueueMaxSize As Long
Private arr() As Long
Private Count As Long

Property Let MaxSize(value As Long)
    QueueMaxSize = value
    ReDim arr(QueueMaxSize - 1) As Long
       
    Head = value
    Tail = 0
    Count = 0
End Property

Function PutIn(Item As Long)
    If Count = QueueMaxSize - 1 Then
        MsgBox "QueueÂú"
        Exit Function
    End If
   
    arr(Tail) = Item
    Tail = (Tail + 1) Mod QueueMaxSize
    Count = Count + 1
End Function

Function GetItem() As Long
    If Count = 0 Then
        MsgBox "¿ÕQueue"
        Exit Function
    End If
   
    Head = Head Mod QueueMaxSize
    GetItem = arr(Head)
    Head = Head + 1
    Count = Count - 1
End Function


Private Sub Class_Initialize()
    Me.MaxSize = 10
End Sub

Private Sub Class_Terminate()
    Erase arr
End Sub

