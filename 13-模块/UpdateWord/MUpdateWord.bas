Attribute VB_Name = "MUpdateWord"
Option Explicit

Enum Pos
    RowStart = 2
    TheName = 2
    LinkWk = 3
    Value = 6
    Cols = 7
End Enum

Type DataStruct
    RowEnd As Long
    Arr() As Variant
    Dic As Object
    
    DocFileName As String
    Doc As Object  'Word.Document
End Type

Sub UpdateWordByCom()  'ͨ����ע��ʽ
    Dim i As Long, k As Long
    Dim Com, ComRngT As String, MyRange, iStart As Long, iEnd As Long
    Dim d As DataStruct
    
    '��ʼ������
    If InitData(d) = -1 Then Exit Sub
    
    If d.Doc.Comments.Count = 0 Then
        MsgBox "��ǰWordû����ע��"
        GoTo A
    End If
    
    For i = 1 To d.Doc.Comments.Count
        Set Com = d.Doc.Comments(i)
        ComRngT = Com.Range.Text
        If d.Dic.Exists(ComRngT) Then
            Set MyRange = Com.Scope
            With MyRange
                iStart = .Start
                iEnd = .End
                iEnd = iEnd - VBA.Len(.Text) + VBA.Len(d.Dic(ComRngT))
                .Text = d.Dic(ComRngT)
            End With
            
            Set MyRange = d.Doc.Range(Start:=iStart, End:=iEnd)
            Com.Delete
            d.Doc.Comments.Add MyRange, ComRngT
        Else
            MsgBox VBA.Format(i, "��0����ע\[") & ComRngT & "]�������Ŀ��Excel��û�С�"
            GoTo A
        End If
    Next i
        
    AppActivate ThisWorkbook.Name
    MsgBox "������ɡ�"
      
A:
'    DocApp.Quit
    Set MyRange = Nothing
    Set d.Doc = Nothing
    Set d.Dic = Nothing
End Sub



