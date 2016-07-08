VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "  二维条码"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpDefaultChar As Any, ByRef lpUsedDefaultChar As Any) As Long
Private Const CP_UTF8 As Long = 65001
Private obj As New clsQRCode

Private Sub CommandButton1_Click()
    Dim b2() As Byte
    Dim s As String
    Dim i As Long, m As Long
'    For i = 0 To cmb1.UBound
'        If cmb1(i).ListIndex < 0 Then Exit Sub
'    Next i
    Select Case ComboBox4.ListIndex
        Case 1
        s = TextBox1.Text
        m = Len(s)
        i = m * 3 + 64
        ReDim b2(i)
        m = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(s), m, b2(0), i, ByVal 0, ByVal 0)
        Case Else
        s = StrConv(TextBox1.Text, vbFromUnicode)
        b2 = s
        m = LenB(s)
    End Select
    Set Image1.Picture = obj.Encode(b2, m, ComboBox1.ListIndex, ComboBox2.ListIndex + 1, ComboBox3.ListIndex - 1)
    SavePicture Image1.Picture, ThisWorkbook.Path & "\" & Format(Now, "YYYY-MM-DD-HH-MM-SS") & ".jpg"
End Sub



Private Sub UserForm_Initialize()
    Dim i As Long
    ComboBox1.AddItem "自动"
    For i = 1 To 40
        ComboBox1.AddItem CStr(i)
    Next i
    ComboBox1.ListIndex = 0
    ComboBox2.AddItem "L - 7%"
    ComboBox2.AddItem "M - 15%"
    ComboBox2.AddItem "Q - 25%"
    ComboBox2.AddItem "H - 30%"
    ComboBox2.ListIndex = 1
    ComboBox3.AddItem "自动"
    For i = 0 To 7
        ComboBox3.AddItem CStr(i)
    Next i
    ComboBox3.ListIndex = 0
    ComboBox4.AddItem "ANSI"
    ComboBox4.AddItem "UTF-8"
    ComboBox4.ListIndex = 1
End Sub
'
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    ThisWorkbook.Close True
'End Sub
