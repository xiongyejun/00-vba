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

Private nTop As CNode

Property Get Top() As CNode
    Set Top = nTop
End Property

'µ¯³ö
Function Pop() As CNode
    If nTop Is Nothing Then
        MsgBox "¿ÕÕ»¡£"
        Exit Function
    End If
    
    Set Pop = nTop
    Set nTop = nTop.Pre
End Function
'Ñ¹Èë
Function Push(n As CNode)
    Set n.Pre = nTop
    Set nTop = n
End Function
'Çå¿Õ
Function EmptyStack()
    Dim tmp As CNode
    
    Do Until nTop Is Nothing
        Set tmp = nTop.Pre
        Set nTop = Nothing
        Set nTop = tmp
    Loop
    Set tmp = Nothing
End Function

Private Sub Class_Terminate()
    EmptyStack
End Sub

