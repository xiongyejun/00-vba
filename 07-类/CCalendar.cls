VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CCalendar 自定义日历控件中day按钮事件
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private Sub btn_Click()
    Dim i_row As Integer, i_col As Integer
    
    '获取当前按钮在ArrC_cal中所在的行号和列号
    i_row = VBA.Val(VBA.Split(btn.Tag, "、")(0))
    i_col = VBA.Val(VBA.Split(btn.Tag, "、")(1))
    With btn.Parent
        .SetBtnBackColor '恢复上1个按钮的背景色
        .Day = VBA.Val(btn.Caption)
        .ArrRow = i_row
        .ArrCol = i_col
    End With
    btn.BackColor = &H8000000C
End Sub

Private Sub btn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btn.Parent.Hide
End Sub

'    Dim f_cld As FCalendar
'
'    Set f_cld = New FCalendar
'
'    f_cld.Show
'
'    ActiveCell.Value = f_cld.TheDate
