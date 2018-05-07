VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FCompare 
   Caption         =   "Compare"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "FCompare.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "FCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frSheet As MSForms.Frame
Private frFieldsCondition As MSForms.Frame
Private frFieldsData As MSForms.Frame

Private frBtn As MSForms.Frame
Private WithEvents btnSelectField As MSForms.CommandButton
Attribute btnSelectField.VB_VarHelpID = -1
Private WithEvents btnOK As MSForms.CommandButton
Attribute btnOK.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1
Private tbSqlOther As MSForms.TextBox '手写一些其他的限制条件

Private Sheet_Names() As String
Private bCancel As Boolean '是否是取消了或者是点了X
Private b_Select As Boolean
Private b_Close As Boolean

Private Const CB_HEIGHT As Integer = 24
Private Const FR_HEIGHT As Integer = 200

'工作表，用来设置选择工作表的checkbox
Property Let SheetNames(Values() As String)
    Sheet_Names = Values
    
    frameAddCheckBox frSheet, Sheet_Names, FR_HEIGHT, , CB_HEIGHT
End Property
Property Get SheetNames() As String()
   SheetNames = GetArrFromFrame(frSheet)
End Property
'表格的字段就以第一个表的为准
Property Get FieldsCondition() As String()
    FieldsCondition = GetArrFromFrame(frFieldsCondition)
End Property
'表格的字段就以第一个表的为准
Property Get FieldsData() As String()
    FieldsData = GetArrFromFrame(frFieldsData)
End Property
Property Get SqlOther() As String
    SqlOther = tbSqlOther.Text
End Property

Property Get Cancel() As Boolean
    Cancel = bCancel
End Property
'是否需要关闭窗体
Property Get bClose() As Boolean
    bClose = b_Close
End Property
'是否选择了sheet或者字段
Property Get bSelect() As Boolean
    bSelect = b_Select
End Property
Property Let bSelect(value As Boolean)
    b_Select = value
End Property


Private Function GetArrFromFrame(fr As MSForms.Frame) As String()
    Dim i As Long, iNext As Long
    Dim arr() As String
    Dim n As Long
    
    n = fr.Controls.Count
    If n Then
        ReDim arr(n - 1) As String
        For i = 0 To n - 1
            If fr.Controls(i).value Then
                arr(iNext) = fr.Controls(i).Caption
                iNext = iNext + 1
            End If
        Next
    End If

    If iNext Then
        ReDim Preserve arr(iNext - 1) As String
        GetArrFromFrame = arr
        b_Select = True
    Else
        b_Select = False
    End If
End Function

Private Sub btnCancel_Click()
    bCancel = True
    b_Close = False
    Me.Hide
End Sub

Private Sub btnOK_Click()
    bCancel = False
    b_Close = False
    Me.Hide
End Sub

Private Sub btnSelectField_Click()
    Dim i As Long, j As Long
    Dim arr() As String
    Dim tmp As Variant
    
    For i = 0 To frSheet.Controls.Count - 1
        If frSheet.Controls(i).value Then
            tmp = MFunc.GetFields(ActiveWorkbook.Worksheets(frSheet.Controls(i).Caption))
            If VBA.IsArray(tmp) Then
                ReDim arr(UBound(tmp)) As String
                For j = 0 To UBound(tmp)
                    arr(j) = VBA.CStr(tmp(j))
                Next
                
                frameAddCheckBox frFieldsCondition, arr, FR_HEIGHT, , CB_HEIGHT
                frameAddCheckBox frFieldsData, arr, FR_HEIGHT, , CB_HEIGHT
            End If
            
            Exit For
        End If
    Next
    
End Sub

Private Sub UserForm_Initialize()
    Dim iLeft As Integer
    iLeft = 5
    Set frSheet = FrameAdd(Me, "选择工作表", iLeft, 5, , FR_HEIGHT): iLeft = iLeft + frSheet.Width
    Set frFieldsCondition = FrameAdd(Me, "选择作为条件的字段", iLeft, 5, , FR_HEIGHT): iLeft = iLeft + frFieldsCondition.Width
    Set frFieldsData = FrameAdd(Me, "选择要对比的字段", iLeft, 5, , FR_HEIGHT): iLeft = iLeft + frFieldsData.Width
    
    
    Set frBtn = FrameAdd(Me, "", 5, FR_HEIGHT + 30, frSheet.Width * 3, 80)
    Set btnSelectField = frBtn.Controls.Add("Forms.CommandButton.1")
    btnSelectField.Caption = "选择字段": btnSelectField.Left = 5
    
    Set btnOK = frBtn.Controls.Add("Forms.CommandButton.1")
    btnOK.Caption = "OK": btnOK.Left = 5 + btnSelectField.Width
    
    Set btnCancel = frBtn.Controls.Add("Forms.CommandButton.1")
    btnCancel.Caption = "取消": btnCancel.Left = 5 + btnSelectField.Width + btnSelectField.Width
    
    '用来填写一些其他的限制条件
    Dim lb As MSForms.Label
    Set lb = frBtn.Controls.Add("Forms.Label.1")
    lb.Caption = "其他的限制条件，如:Where 专业<>'不计价'"
    lb.Top = btnOK.Top + btnOK.Height + 10
    lb.Width = frBtn.Width
    
    Set tbSqlOther = frBtn.Controls.Add("Forms.TextBox.1")
    tbSqlOther.Top = lb.Top + lb.Height - 5
    tbSqlOther.Width = frBtn.Width
    tbSqlOther.Height = 40
    tbSqlOther.Text = "Where 专业<>'不计价'"
    
    Me.Height = frBtn.Top + frBtn.Height + 30
    Me.Width = iLeft + 20
    Me.Caption = "Compare--表格的字段就以第一个表的为准，假定各表字段一模一样。"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If 0 = CloseMode Then
        bCancel = True
        b_Close = True
    Else
        bCancel = False
    End If
    
    Cancel = True
    Me.Hide
End Sub
