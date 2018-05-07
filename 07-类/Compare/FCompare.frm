VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FCompare 
   Caption         =   "Compare"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "FCompare.frx":0000
   StartUpPosition =   1  '����������
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
Private tbSqlOther As MSForms.TextBox '��дһЩ��������������

Private Sheet_Names() As String
Private bCancel As Boolean '�Ƿ���ȡ���˻����ǵ���X
Private b_Select As Boolean
Private b_Close As Boolean

Private Const CB_HEIGHT As Integer = 24
Private Const FR_HEIGHT As Integer = 200

'��������������ѡ�������checkbox
Property Let SheetNames(Values() As String)
    Sheet_Names = Values
    
    frameAddCheckBox frSheet, Sheet_Names, FR_HEIGHT, , CB_HEIGHT
End Property
Property Get SheetNames() As String()
   SheetNames = GetArrFromFrame(frSheet)
End Property
'�����ֶξ��Ե�һ�����Ϊ׼
Property Get FieldsCondition() As String()
    FieldsCondition = GetArrFromFrame(frFieldsCondition)
End Property
'�����ֶξ��Ե�һ�����Ϊ׼
Property Get FieldsData() As String()
    FieldsData = GetArrFromFrame(frFieldsData)
End Property
Property Get SqlOther() As String
    SqlOther = tbSqlOther.Text
End Property

Property Get Cancel() As Boolean
    Cancel = bCancel
End Property
'�Ƿ���Ҫ�رմ���
Property Get bClose() As Boolean
    bClose = b_Close
End Property
'�Ƿ�ѡ����sheet�����ֶ�
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
    Set frSheet = FrameAdd(Me, "ѡ������", iLeft, 5, , FR_HEIGHT): iLeft = iLeft + frSheet.Width
    Set frFieldsCondition = FrameAdd(Me, "ѡ����Ϊ�������ֶ�", iLeft, 5, , FR_HEIGHT): iLeft = iLeft + frFieldsCondition.Width
    Set frFieldsData = FrameAdd(Me, "ѡ��Ҫ�Աȵ��ֶ�", iLeft, 5, , FR_HEIGHT): iLeft = iLeft + frFieldsData.Width
    
    
    Set frBtn = FrameAdd(Me, "", 5, FR_HEIGHT + 30, frSheet.Width * 3, 80)
    Set btnSelectField = frBtn.Controls.Add("Forms.CommandButton.1")
    btnSelectField.Caption = "ѡ���ֶ�": btnSelectField.Left = 5
    
    Set btnOK = frBtn.Controls.Add("Forms.CommandButton.1")
    btnOK.Caption = "OK": btnOK.Left = 5 + btnSelectField.Width
    
    Set btnCancel = frBtn.Controls.Add("Forms.CommandButton.1")
    btnCancel.Caption = "ȡ��": btnCancel.Left = 5 + btnSelectField.Width + btnSelectField.Width
    
    '������дһЩ��������������
    Dim lb As MSForms.Label
    Set lb = frBtn.Controls.Add("Forms.Label.1")
    lb.Caption = "������������������:Where רҵ<>'���Ƽ�'"
    lb.Top = btnOK.Top + btnOK.Height + 10
    lb.Width = frBtn.Width
    
    Set tbSqlOther = frBtn.Controls.Add("Forms.TextBox.1")
    tbSqlOther.Top = lb.Top + lb.Height - 5
    tbSqlOther.Width = frBtn.Width
    tbSqlOther.Height = 40
    tbSqlOther.Text = "Where רҵ<>'���Ƽ�'"
    
    Me.Height = frBtn.Top + frBtn.Height + 30
    Me.Width = iLeft + 20
    Me.Caption = "Compare--�����ֶξ��Ե�һ�����Ϊ׼���ٶ������ֶ�һģһ����"
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
