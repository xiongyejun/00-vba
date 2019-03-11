Attribute VB_Name = "MUpdateExcel"
Option Explicit

Sub UpdateExcel()
    Dim d As Object '��ȡ���ظ��Ĺ���������
    Set d = CreateObject("Scripting.Dictionary") '�����ֵ���󣬺��ڰ󶨣�����Ҫ�����ã����ߡ����á������C:\WINDOWS\system32\scrrun.dll)
    
    Dim irow As Long, i As Long
    Dim wk As Workbook, Path As String, sht As Worksheet
    Dim Arr(), strKey As String

    Set wk = ActiveWorkbook
    Set sht = ActiveSheet
    Path = wk.Path & "\"
    irow = Cells(Cells.Rows.Count, Pos.LinkWk).End(xlUp).Row
    If irow < Pos.RowStart Then Exit Sub
    Arr = Range("A1").Resize(irow, Pos.LinkRng).Value
    
    '��ȡ���еĹ���������
    For i = Pos.RowStart To irow
        strKey = VBA.CStr(Arr(i, Pos.LinkWk))
        If VBA.Len(strKey) Then d(strKey) = i
    Next i
    
    Dim ArrWk
    ArrWk = d.Keys()
    '�����еĹ�����
    For i = 0 To d.Count - 1
        Workbooks.Open Path & VBA.CStr(ArrWk(i)), UpdateLinks:=False
    Next i
    
    Dim strWk As String, strSht As String, strRng As String
    '��������
    ActiveWorkbook.Activate
    For i = Pos.RowStart To irow
        strWk = VBA.CStr(Arr(i, Pos.LinkWk))
        If VBA.Len(strWk) Then
            strSht = VBA.CStr(Arr(i, Pos.LinkSheet))
            If VBA.Len(strSht) Then
                strRng = VBA.CStr(Arr(i, Pos.LinkRng))
                If VBA.Len(strRng) Then
                    '��������--��Ϊ�е��ǹ�ʽ����ģ��Ḳ��
                    sht.Cells(i, Pos.Value).Value = Workbooks(strWk).Worksheets(strSht).Range(strRng).Value
                End If
            End If
        End If
    Next i
    wk.Activate
    
A:
    For i = 0 To d.Count - 1
        Workbooks(d.Keys()(i)).Close False
    Next i
    MsgBox "�������"
    
    Set wk = Nothing
    Set d = Nothing
End Sub
