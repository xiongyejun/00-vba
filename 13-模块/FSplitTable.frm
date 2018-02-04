VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FSplitTable 
   Caption         =   "UserForm2"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FSplitTable.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "FSplitTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�����ֵ��¼ÿһ�����͵�rng��Χ
'Ȼ��һ���Ը��ƹ�ȥ

Dim lb_head As MSForms.Label
Dim WithEvents tb_head As MSForms.TextBox
Attribute tb_head.VB_VarHelpID = -1
Dim WithEvents btn_split As MSForms.CommandButton
Attribute btn_split.VB_VarHelpID = -1
Dim cb_saveBook As MSForms.CheckBox

Dim fr_col As MSForms.Frame
Dim rng_head As Range


Private Sub btn_split_Click()
    Dim arr_col() As Long
    Dim dic As Object
    Dim i_row As Long, i As Long, j As Long
    Dim str_key As String
    Dim k_col As Long
    Dim arr()
    Dim path As String
    Dim i_cols As Long  '�����еķ�Χ����k_col��һ��һ����k_col�ǲ��ظ���
    
    On Error GoTo Err1
    
    k_col = GetCol(arr_col) '��ȡ�б����Ӧ���к�arr_col
    If k_col = 0 Then Exit Sub
    
    ActiveSheet.AutoFilterMode = False
    '��λ���ݷ�Χ
    i_row = Cells(Cells.Rows.Count, rng_head.Column).End(xlUp).Row
    i_cols = rng_head.Columns.Count
    If i_row < 2 Then MsgBox "û������": Exit Sub
    arr = Range(Cells(1, rng_head.Column), Cells(i_row, rng_head.Column + i_cols - 1)).Value
    
    '���ֵ��¼ÿһ�����͵�rng��Χ
    Set dic = CreateObject("Scripting.Dictionary")
    For i = rng_head.Row + 1 To i_row
        str_key = ""
        For j = 0 To k_col - 1
            str_key = str_key & "|" & arr(i, arr_col(j))
        Next j
        
        If dic.exists(str_key) Then
            Set dic(str_key) = Excel.Union(dic(str_key), Cells(i, rng_head.Column).Resize(1, i_cols))
        Else
            Set dic(str_key) = rng_head
            Set dic(str_key) = Excel.Union(dic(str_key), Cells(i, rng_head.Column).Resize(1, i_cols))
        End If
    Next i
    
    '��������
    path = ActiveWorkbook.path & "\"
    For i = 0 To dic.Count - 1
        str_key = dic.Keys()(i)
        If cb_saveBook.Value Then
            Workbooks.Add
            ActiveWorkbook.SaveAs path & VBA.Replace(str_key, "|", "")
            'ɾ�������sheet��ֻ����1��
            Application.DisplayAlerts = False
            For j = Worksheets.Count To 2 Step -1
                Worksheets(j).Delete
            Next
            Application.DisplayAlerts = True
            
            dic(str_key).Copy Range("A1")
            ActiveWorkbook.Close True
        Else
            Worksheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = str_key
            dic(str_key).Copy Range("A1")
        End If
        
    Next i
    
Err1:
    If Err.Number <> 0 Then MsgBox Err.Description
    
    Unload Me
End Sub
'��ȡÿ���б������ڵ��к�
'�����еĲ��ظ�����
Function GetCol(arr_col() As Long) As Long
    Dim ct As Control
    Dim arr()
    Dim dic As Object
    Dim k As Long
    
    arr = rng_head.Value
    arr = Application.WorksheetFunction.Transpose(arr)
    Func.DataToDic dic, arr, 1 '��¼�б������ڵ��к�
    
    k = 0
    For Each ct In fr_col.Controls
        If ct.Value <> "" Then
            ReDim Preserve arr_col(k) As Long
            arr_col(k) = dic(ct.Value) + rng_head.Column - 1 '��¼ʵ�����ڵ��кţ���Ϊ�����п��ܲ���A�п�ʼ
            k = k + 1
        End If
    Next
    GetCol = k
    
    Set dic = Nothing
End Function
'ѡ���б���ķ�Χ
Private Sub tb_head_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim end_col As Long
    
    Set rng_head = Nothing
    
    end_col = Range("A1").End(xlToRight).Column
    If end_col > 30 Then end_col = 30
    On Error Resume Next
    Set rng_head = Application.InputBox("ѡ�����������", Default:=Range("A1").Resize(1, end_col).Address, Type:=8)
    On Error GoTo 0
    
    If Not rng_head Is Nothing Then
        tb_head.Text = rng_head.Address
        FrameAdd
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim i_left As Integer
    Dim i_top As Integer
    
    i_left = 5
    i_top = 5
    
    Set lb_head = Func.labelAdd(Me, "����", i_left, i_top + 5, 20)
    i_left = lb_head.Width + i_left
    Set tb_head = Func.tbAdd(Me, "", i_left, i_top, 150)
    i_left = tb_head.Width + i_left
    Set btn_split = Func.btnAdd(Me, "���", i_left, i_top)
    i_left = btn_split.Width + i_left
    Set cb_saveBook = Func.cbAdd(Me, "��湤����", i_left, i_top)
    
    
    i_left = 5
    i_top = i_top + tb_head.Height + 10
    Set fr_col = Func.FrameAdd(Me, "ѡ����", i_left, i_top, 300)
    
    Me.Width = fr_col.Width + 20
End Sub

Function FrameAdd()
    Dim cb As MSForms.ComboBox
    Dim i As Long
    Dim arr() As String
    Dim i_top As Integer
    
    i_top = 5
    
    ReDim arr(rng_head.Columns.Count - 1) As String
    For i = 1 To rng_head.Columns.Count
        arr(i - 1) = VBA.CStr(rng_head.Cells(1, i).Value)
    Next i
    
    fr_col.Controls.Clear
    For i = 1 To rng_head.Columns.Count
        Set cb = Func.ComboBoxAdd(fr_col, 5, i_top, fr_col.Width - 20)
        i_top = i_top + cb.Height + 5
        cb.List = arr
    Next i
    
    If i_top > 300 Then
        fr_col.Height = 300 + 10
        fr_col.ScrollBars = fmScrollBarsVertical
        fr_col.ScrollHeight = i_top
    Else
        fr_col.Height = i_top + 10
    End If
    
    Me.Height = fr_col.Height + 30 + fr_col.Top
End Function
