Attribute VB_Name = "M�������"
Option Explicit

'��i_row����������ѡ��length�����������

Type PassWord
    arr_psd() As String
    arr_p_key() As Integer      'ָ��arr_key���±�
    num As Integer              '��ǰ�������±�
    length As Integer            '���볤��
    str As String
End Type

Sub vba_main()
    Dim arr_key()
    Dim i_row As Long
    Dim i As Long
    Dim pw As PassWord
    Dim tmp_num As Integer
'    Dim rng As Range
    Dim i_len As Integer
    Dim wk As Workbook
    Dim k As Long, k_sum As Long
    Dim file_name As String
   
    file_name = GetFileName()
    If file_name = "" Then Exit Sub
'    Set rng = Range("C1")
'    Range("C:C").Clear
   
    ActiveSheet.AutoFilterMode = False
    i_row = Range("A" & Cells.Rows.Count).End(xlUp).Row
    ReDim arr_key(i_row - 2)
    For i = 2 To i_row
        arr_key(i - 2) = Cells(i, 1).Value
    Next i
   
    i_row = i_row - 2   'key���ϱ꣬�±���0��arr_psd�ĸ���-1
   
'    pw.length = Range("B2").Value
    On Error Resume Next
   
    k = 0
    For i_len = 1 To i_row + 1
        pw.length = i_len
        k_sum = (i_row + 1) ^ i_len
       
        InitPwType pw, VBA.CStr(arr_key(0))
       
        Do Until pw.num > pw.length  '���仯���Ǹ�λ�ò��ܳ�������ĳ���
           
            Do Until tmp_num >= pw.num      '��λ��λ�ò��ܳ����仯��λ��
                For i = 0 To i_row
                    k = k + 1
                    pw.arr_psd(0) = arr_key(i)
                    pw.str = VBA.Join(pw.arr_psd, "")
    '                rng.Value = "'" & pw.str
    '                Set rng = rng.Offset(1, 0)
                   
                    Application.StatusBar = "���ڲ���" & i_len & "λ���룺" & k & "/" & k_sum & "  " & pw.str
                    Set wk = Workbooks.Open(Filename:=file_name, UpdateLinks:=False, PassWord:=pw.str)
                    If Not wk Is Nothing Then
                        Debug.Print pw.str
                        Exit Sub
                    End If
                   
                Next i
                pw.arr_p_key(0) = i_row + 1
               
                tmp_num = 0
                '�е����Ƽӷ���10��1
                Do
                    'Ҫ��λ��
                    pw.arr_psd(tmp_num) = arr_key(0)
                    pw.arr_p_key(tmp_num) = 0
                   
                    tmp_num = tmp_num + 1
                    If tmp_num = pw.length Then Exit Do
                   
                    pw.arr_p_key(tmp_num) = pw.arr_p_key(tmp_num) + 1
                   
                    If pw.arr_p_key(tmp_num) <= i_row Then  '����key���±�λ��
                        pw.arr_psd(tmp_num) = arr_key(pw.arr_p_key(tmp_num))
                    End If
                Loop While pw.arr_p_key(tmp_num) = i_row + 1 '������key���±꣬������λ
               
            Loop
           
            pw.num = pw.num + 1
        Loop
       
    Next i_len
   
    On Error GoTo err_handle
   
   
    Exit Sub
   
err_handle:
    MsgBox Err.Description
End Sub

Function InitPwType(pw As PassWord, first_key As String)
    Dim i As Long
   
    ReDim pw.arr_psd(pw.length - 1) As String
    ReDim pw.arr_p_key(pw.length - 1) As Integer
    pw.num = 0
   
    For i = 0 To pw.length - 1
        pw.arr_psd(i) = first_key
    Next i
End Function

Function GetFileName() As String
    With Application.FileDialog(msoFileDialogOpen)
        .InitialFileName = ActiveWorkbook.Path & "\"
        .Filters.Clear
'        .Filters.Add "CSV TXT", "*.csv;*.txt"
        
        If .Show = -1 Then                  ' -1����ȷ����0����ȡ��
            GetFileName = .SelectedItems(1)
        Else
            GetFileName = ""
            MsgBox "��ѡ���ļ�����"
        End If
    End With
End Function

