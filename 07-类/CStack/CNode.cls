VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Node
    Value As Long
    Pre As CNode
    pNext As CNode
    Height As Long  '
End Type
Private n As Node

Property Let Value(v As Long)
    n.Value = v
End Property
Property Get Value() As Long
    Value = n.Value
End Property

Property Let Height(v As Long)
    n.Height = v
End Property
Property Get Height() As Long
    Height = n.Height
End Property

Property Set Pre(v As CNode)
    Set n.Pre = v
End Property
Property Get Pre() As CNode
    Set Pre = n.Pre
End Property

Property Set pNext(v As CNode)
    Set n.pNext = v
End Property
Property Get pNext() As CNode
    Set pNext = n.pNext
End Property


