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
    Left As CNode
    Right As CNode
    N As Long
End Type
Private nd As Node

Property Let Value(v As Long)
    nd.Value = v
End Property
Property Get Value() As Long
    Value = nd.Value
End Property

Property Let N(v As Long)
    nd.N = v
End Property
Property Get N() As Long
    N = nd.N
End Property

Property Set Left(v As CNode)
    Set nd.Left = v
End Property
Property Get Left() As CNode
    Set Left = nd.Left
End Property

Property Set Right(v As CNode)
    Set nd.Right = v
End Property
Property Get Right() As CNode
    Set Right = nd.Right
End Property

