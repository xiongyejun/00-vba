Attribute VB_Name = "MUpdateExcel"
Option Explicit

Sub UpdateExcel()
    shtLink.Activate

    Dim d As Object '��ȡ���ظ��Ĺ���������
    Set d = CreateObject("Scripting.Dictionary") '�����ֵ���󣬺��ڰ󶨣�����Ҫ�����ã����ߡ����á������C:\WINDOWS\system32\scrrun.dll)
    
    Dim iRow As Long, i As Long
    Dim wk As Workbook, path As String
    
    path = ThisWorkbook.path & "\"
    iRow = Cells(Cells.Rows.Count, Pos.LinkWk).End(xlUp).Row
    '��ȡ���еĹ���������
    For i = 3 To iRow
        If Cells(i, "C").Value <> "" Then d(Cells(i, "C").Value) = ""
    Next i
    
    '�����еĹ�����
    For i = 0 To d.Count - 1
        Workbooks.Open path & d.Keys()(i), UpdateLinks:=False
    Next i
    
    ThisWorkbook.Activate
    For i = 3 To iRow
        If Cells(i, "E").Value <> "" Then
            Set wk = Workbooks(Cells(i, "C").Value)
            Cells(i, "F").Value = wk.Worksheets(Cells(i, "D").Value).Range(Cells(i, "E").Value).Value
        End If
    Next i
    
A:
    For i = 0 To d.Count - 1
       Workbooks(d.Keys()(i)).Close False
    Next i
    MsgBox "�������"
    
    Set wk = Nothing
    Set d = Nothing

End Sub
