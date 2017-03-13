VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfSql 
   Caption         =   "UserForm2"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "usfSql.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "usfSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�򿪴��壬��ȡ��ǰ��������
'��hide�Ա�֤����
'�༭��sql�����ִ��

Private lbSoure As MSForms.Label                            '����Դ��ǩ
Attribute lbSoure.VB_VarHelpID = -1
Private WithEvents tbSoure As MSForms.TextBox               '����Դ�ı���
Attribute tbSoure.VB_VarHelpID = -1
Private WithEvents btnSoure As MSForms.CommandButton        '����Դ��ť
Attribute btnSoure.VB_VarHelpID = -1

Private lbTable As MSForms.Label                            '�������ǩ
Private WithEvents cbTable As MSForms.ComboBox              '�������б�
Attribute cbTable.VB_VarHelpID = -1

Private WithEvents frField As MSForms.Frame                            '�ֶο�
Attribute frField.VB_VarHelpID = -1

Private frSQL As MSForms.Frame                              'sql
Private tbSQL As MSForms.TextBox                            'SQL�ı���
Private WithEvents btnGetField As MSForms.CommandButton     '��ȡ�ֶ�
Attribute btnGetField.VB_VarHelpID = -1
Private WithEvents btnExecute As MSForms.CommandButton      'ִ�а�ť
Attribute btnExecute.VB_VarHelpID = -1
Private WithEvents btnHideForm As MSForms.CommandButton     '���ش��尴ť
Attribute btnHideForm.VB_VarHelpID = -1

Private WithEvents lbSqlOrder As MSForms.listBox            '�г����
Attribute lbSqlOrder.VB_VarHelpID = -1


Private Sub btnExecute_Click()
    Dim rng As Range
    Dim iAdo As Long
    
    Me.Hide                 '���ش���
    
    Func.getRngByInputBox rng
    
    If Not rng Is Nothing Then
        On Error Resume Next
        rng.Comment.Delete
        On Error GoTo 0
                
        iAdo = Func.CreateAdo(tbSQL.Text, rng, tbSoure.Text)
        If iAdo = 1 Then
            rng.AddComment
            rng.Comment.Text Text:=tbSoure.Text & vbNewLine & tbSQL.Text         '��sql���ŵ���ע��
        Else
            Me.Show
        End If
    End If

    Set rng = Nothing
    Exit Sub
End Sub

Private Sub btnGetField_Click()
    Dim ctl As Control, strField As String
    Dim str As String, i As Long
    Dim Arr() As String, k As Long
    
    k = 0
    For Each ctl In frField.Controls
        If ctl.value Then
            k = k + 1
            ReDim Preserve Arr(1 To k) As String
            Arr(k) = ctl.Caption
        End If
    Next ctl
    strField = "Select [" & Join(Arr, "],[") & "] From "
    str = tbSQL.Text
    i = InStr(str, "From")
    If i > 0 Then str = Right(str, Len(str) - i - 4)
    tbSQL.Text = strField & str
    
    Set ctl = Nothing
    Erase Arr
End Sub

Private Sub btnHideForm_Click()
    Me.Hide
End Sub

Private Sub btnSoure_Click()        '��ȡ�ļ�����
    Dim str As String
    str = GetFileName
    If str <> "" Then
        tbSoure.Text = str
    End If
End Sub

Private Sub cbTable_Change()
    '��ȡ�ֶΣ���ӿؼ�
    Dim shtName As String
    Dim iTop As Long        '��ӿؼ����ֶ�frField����ײ�
    
    On Error GoTo Err
    
    If cbTable.Text = "" Then Exit Sub
    
    shtName = Split(cbTable.Text, "$")(0)
    
    iTop = getField(tbSoure.Text, shtName)
    '��������lbSQL����tbSQL
    iTop = iTop + 5
    frSQL.Top = iTop
    lbSqlOrder.Height = iTop + frSQL.Height
    
    Me.Height = iTop + frSQL.Height + 30
    
'    Me.Left = Func.setFormPosLeft(Me.Width)
'    Me.Top = Func.setFormPosRight(Me.Height)
'    Me.Left = (ActiveWindow.Width - Me.Width) / 2
'    Me.Top = (ActiveWindow.Height - Me.Height) / 2
    
    Dim strSql As String
    Dim strTabel As String
    Dim strFrom As String
    
    strTabel = "From [" & cbTable.Text & "]"
    strSql = tbSQL.Text
    strFrom = Split(strSql, "From ")(1)
    If strFrom <> "" Then
        strFrom = Split(strFrom, "]")(0)
        strFrom = strFrom & "]"
    End If
    
    strFrom = "From " & strFrom
    
    strSql = Replace(strSql, strFrom, strTabel)
    tbSQL.Text = strSql & " "
    Exit Sub
    
Err:
    MsgBox Err.Description
End Sub

Private Sub frField_Click()
    Func.frameCheckBoxValue frField
End Sub

Private Sub frField_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Func.frameCheckBoxValue frField, False
End Sub

Private Sub lbSqlOrder_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    tbSQL.Text = tbSQL.Text & lbSqlOrder.value
End Sub

Private Sub tbSoure_Change()
    getShtName tbSoure.Text         '��ȡ��������������
End Sub

Private Sub tbSoure_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo Err
    Workbooks.Open (tbSoure.Text)
    Exit Sub
Err:
    MsgBox Err.Description
End Sub

Private Sub UserForm_Initialize()
    Dim iLeft As Integer, iTop As Integer
    Dim cbArr(1 To 20, 1 To 2) As String                '��ѯ���ѡ��where��order by ����
    Dim iCb As Long, i As Long
    
    Const lbWidth As Long = 30
    Const tbWidth As Long = 80 * 4
    Const btnWidth As Long = 80
    Const btnHeight As Long = 24
    Const iStep As Long = 5

    '�������Դ��ǩ���ı���
    iLeft = 2
    Set lbSoure = labelAdd(Me, "����Դ", iLeft, iTop + 4, lbWidth)
    iLeft = iLeft + lbSoure.Width
    
    Set tbSoure = tbAdd(Me, ActiveWorkbook.fullName, iLeft, iTop, tbWidth - 30 - lbWidth, 30)
    tbSoure.MultiLine = True
    
    Set btnSoure = btnAdd(Me, "���", iLeft + tbSoure.Width, iTop, 30, 30)
    iTop = iTop + tbSoure.Height
     
    '��ӹ������ǩ��ComboBox
    iLeft = 2
    iTop = iTop + iStep
    Set lbTable = labelAdd(Me, "������", iLeft, iTop + 4, lbWidth)
    iLeft = iLeft + lbTable.Width
    
    Set cbTable = ComboBoxAdd(Me, iLeft, iTop, tbWidth - lbWidth) '��ӹ������б��
    
    iTop = iTop + cbTable.Height
    getShtName ActiveWorkbook.fullName                            '��ȡ��������������
    
    iLeft = 2
    iTop = iTop + iStep
    '����ֶεĿ��frField
    Set frField = FrameAdd(Me, "�ֶ�", iLeft, iTop, tbWidth)
    iTop = iTop + frField.Height
    
    iLeft = 2
    Set frSQL = FrameAdd(Me, "SQL���", iLeft, iTop + 4, tbWidth)
    frSQL.ForeColor = &H8000000D
    
    Set tbSQL = tbAdd(frSQL, "Select * From ", iLeft, 5, tbWidth - 8, 50)
    iTop = tbSQL.Top + tbSQL.Height
    Set btnGetField = btnAdd(frSQL, "��ȡ�ֶ�", iLeft, iTop, btnWidth)
    iLeft = iLeft + btnGetField.Width
    Set btnExecute = btnAdd(frSQL, "ִ��SQL", iLeft, iTop, btnWidth)
    iLeft = iLeft + btnExecute.Width
    Set btnHideForm = btnAdd(frSQL, "���ش���", iLeft, iTop, btnWidth)
    btnHideForm.Cancel = True
    
    frSQL.Height = iTop + btnExecute.Height + 10
    tbSQL.MultiLine = True
    
    '���õ�SQL�ؼ���
    Set lbSqlOrder = listBoxAdd(Me, tbWidth + 5, 0, 100, frSQL.Top + frSQL.Height)
    lbSqlOrderAddItem
       
    With Me
        .Width = tbWidth + lbSqlOrder.Width + 12
        .Caption = "Ado��ѯ"
        .Height = frSQL.Top + frSQL.Height + 30
    End With
    
End Sub

'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    Me.Hide
''    Cancel = 1
'End Sub

Private Function getShtName(strWk As String)
    Dim wk As Workbook
    Dim shtNameArr() As String                  '����������
    Dim ifOpen As Long
    
    Application.ScreenUpdating = False
    ifOpen = getWorkbook(wk, strWk)
    
    If ifOpen > 0 Then
        Func.getShtNameFromWorkbook wk, shtNameArr, True
        cbTable.Clear
        cbTable.List = shtNameArr
        cbTable.RemoveItem 0
        
        If ifOpen = 1 Then wk.Close False
    End If
    
    Set wk = Nothing
    Erase shtNameArr
    Application.ScreenUpdating = True
End Function

Private Function getField(strWk As String, shtName As String) As Long  '��ȡ�ֶΣ�������ӿؼ����ֶ�frField����ײ�
    Dim wk As Workbook, sht As Worksheet
    Dim ctlArr()                   '�ֶ�
    Dim ifOpen As Long
    Dim i As Long
    
    Application.ScreenUpdating = False
    ifOpen = getWorkbook(wk, strWk)
    
    If ifOpen > 0 Then
        Set sht = wk.Worksheets(shtName)
        sht.Activate
        ReDim ctlArr(Range("IV1").End(xlToLeft).Column - 1)
        
        For i = 0 To Range("IV1").End(xlToLeft).Column - 1
            ctlArr(i) = Cells(1, i + 1).value
            ctlArr(i) = Replace(ctlArr(i), Chr(10), "_")
            ctlArr(i) = Replace(ctlArr(i), ".", "#")
        Next i
        
        frameAddCheckBox frField, ctlArr              '���
        getField = frField.Top + frField.Height
        
        If ifOpen = 1 Then wk.Close False
    End If
    
    Set wk = Nothing
    Set sht = Nothing
    Erase ctlArr
    Application.ScreenUpdating = True
End Function

Function lbSqlOrderAddItem()
    
    With lbSqlOrder
        .AddItem "Distinct "
        .AddItem "Top n "
        .AddItem "Top n Percent "
        .AddItem "Where "
        .AddItem "Where [] In "
        .AddItem "Where [] Between val1 And val2 "
        .AddItem "Group By "
        .AddItem "Group By Having "
        .AddItem "Transform [] Group By [] Pivot []"
        .AddItem "Order By Asc "
        .AddItem "Order By Desc "
        .AddItem "Switch "
        .AddItem "Union All "
        .AddItem "A Inner Join B On compopr"
        .AddItem "A Left Outer Join B On compopr"
        .AddItem "A Right Outer Join B On compopr"
        
        .AddItem "Instr(str,strmatch)"
        .AddItem "Instrrev(str,strmatch)"
'        .AddItem "Strreverse(��ת)"
        .AddItem "Strconv(ת��,conversion)"
        .AddItem "Partition(field,1,100,1)"
    End With
End Function
