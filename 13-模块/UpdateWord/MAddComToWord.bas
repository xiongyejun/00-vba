Attribute VB_Name = "MAddComToWord"
Option Explicit

Sub AddComToWord()
    Dim d As DataStruct
    Dim i As Long
    
    If InitData(d) = -1 Then Exit Sub
    
    '已经有的项目就不需要了
    Dim str_key As String
    For i = 1 To d.Doc.Comments.Count
        str_key = d.Doc.Comments(i).Range.Text
        If d.Dic.Exists(str_key) Then
            d.Dic.Remove str_key
        End If
    Next
    
    Dim tmpKey
    tmpKey = d.Dic.Keys()
    
    Dim p As Object '段落
    For i = 0 To UBound(tmpKey)
        Set p = d.Doc.Paragraphs.Add
        p.Range.Text = tmpKey(i) & vbNewLine
        d.Doc.Comments.Add p.Range, VBA.CStr(tmpKey(i))
    Next
    
    MsgBox VBA.Format(i, "已添加0条批注。")
End Sub
