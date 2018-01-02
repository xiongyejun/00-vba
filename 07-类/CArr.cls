VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CArr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iLen As Long
Private iNext As Long
Private Arr() As String 'item type

Property Get Count() As Long
    Count = iNext
End Property

Property Get Items() As String() 'item type
    If GetItems() Then Items = Arr
End Property

Function Item(i As Long) As String
    If i < iNext Then
        Item = Arr(i)
    End If
End Function

Function Add(Value As String) 'item type)
    If iNext = iLen Then
        RedimArr
    End If
    
    Arr(iNext) = Value
    iNext = iNext + 1
End Function

Function Join(delimiter As String) As String
    If GetItems() Then Join = VBA.Join(Arr, delimiter)
End Function

Private Function GetItems() As Boolean
    'iNextÒª´óÓÚ0
    If iNext Then
        ReDim Preserve Arr(iNext - 1) As String 'item type
        iLen = iNext - 1
        GetItems = True
    Else
        GetItems = False
    End If
End Function


Private Function RedimArr()
    If iLen < 1024 Then
        iLen = iLen * 1.2
    Else
        iLen = iLen * 2
    End If
    
    ReDim Preserve Arr(iLen - 1) As String 'item type
End Function

Private Sub Class_Initialize()
    iLen = 10
    iNext = 0
    ReDim Arr(iLen - 1) As String 'item type
End Sub
