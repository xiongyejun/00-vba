Attribute VB_Name = "MUpdateExcel"
Option Explicit

Sub UpdateExcel()
    shtLink.Activate

    Dim d As Object '获取不重复的工作簿名称
    Set d = CreateObject("Scripting.Dictionary") '创建字典对象，后期绑定，不需要先引用（工具→引用→浏览→C:\WINDOWS\system32\scrrun.dll)
    
    Dim iRow As Long, i As Long
    Dim wk As Workbook, path As String
    
    path = ThisWorkbook.path & "\"
    iRow = Cells(Cells.Rows.Count, Pos.LinkWk).End(xlUp).Row
    '读取所有的工作簿名称
    For i = 3 To iRow
        If Cells(i, "C").Value <> "" Then d(Cells(i, "C").Value) = ""
    Next i
    
    '打开所有的工作簿
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
    MsgBox "更新完成"
    
    Set wk = Nothing
    Set d = Nothing

End Sub
