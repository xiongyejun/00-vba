Attribute VB_Name = "MUpdateWord"
Option Explicit

Enum Pos
    RowStart = 2
    TheName = 2
    LinkWk = 3
    LinkSheet
    LinkRng
    Value
    Unit '���ݵĵ�λ
    PPTValue
    
    Cols = 10
End Enum

Type DataStructDoc
    RowEnd As Long
    Arr() As Variant
    dic As Object
    
    DocFileName As String
    Doc As Object  'Word.Document
End Type

Sub UpdateWordByCom()  'ͨ����ע��ʽ
    Dim i As Long, k As Long
    Dim Com, ComRngT As String, MyRange, iStart As Long, iEnd As Long
    Dim d As DataStructDoc
    
    '��ʼ������
    If InitData(d) = -1 Then Exit Sub
    
    If d.Doc.Comments.Count = 0 Then
        MsgBox "��ǰWordû����ע��"
        GoTo A
    End If
    
    Dim strValue As String
    For i = 1 To d.Doc.Comments.Count
        Set Com = d.Doc.Comments(i)
        ComRngT = Com.Range.Text
        strValue = d.dic(ComRngT)
                    
        If d.dic.Exists(ComRngT) Then
            Set MyRange = Com.Scope
            With MyRange
                iStart = .Start
                iEnd = .End
                iEnd = iEnd - VBA.Len(.Text) + VBA.Len(strValue)
                .Text = strValue
            End With
            
            Set MyRange = d.Doc.Range(Start:=iStart, End:=iEnd)
            Com.Delete
            d.Doc.Comments.Add MyRange, ComRngT
        Else
            If MsgBox(VBA.Format(i, "��0����ע\[") & ComRngT & "]�������Ŀ��Excel��û�С�" & vbNewLine & "�Ƿ������", vbYesNo) = vbNo Then
                GoTo A
            End If
            
        End If

    Next i
        
    On Error Resume Next
    AppActivate ActiveWorkbook.Name
    On Error GoTo 0
    
    MsgBox "������ɡ�"
      
A:
'    DocApp.Quit
    Set MyRange = Nothing
    Set d.Doc = Nothing
    Set d.dic = Nothing
End Sub

Private Function OpenDocument(FilePath As String) As Object
    Dim DocApp As Object 'New Word.Application
    Dim Doc As Object  'Word.Document
    
    On Error Resume Next
    Set DocApp = VBA.GetObject(, "Word.Application")
    If DocApp Is Nothing Then
        Set DocApp = VBA.CreateObject("Word.Application")
    End If
    On Error GoTo 0
    DocApp.Visible = True
    
    Dim FileName As String
    FileName = VBA.Right$(FilePath, VBA.Len(FilePath) - VBA.InStrRev(FilePath, "\"))
    
    '����򿪵�word����fileName�ģ���ʹ�����
    On Error Resume Next
    DocApp.Documents(FileName).Activate
    If Err.Number <> 0 Then
        Set Doc = DocApp.Documents.Open(FilePath)
    Else
        Set Doc = DocApp.Documents(FileName)
        If Doc.FullName <> FilePath Then
            MsgBox "����ͬ���ĵ����ˡ�"
            Set Doc = Nothing
        End If
    End If
    On Error GoTo 0
    
    Set OpenDocument = Doc
End Function

Private Function InitData(d As DataStructDoc) As ReturnCode
    'ѡ��word
    d.DocFileName = GetFileName("doc")
    If VBA.Len(d.DocFileName) = 0 Then
        InitData = -1
        Exit Function
    End If
    
    '��ȡ����
    ActiveSheet.AutoFilterMode = False
    d.RowEnd = Cells(Cells.Rows.Count, Pos.TheName).End(xlUp).Row
    If d.RowEnd < 2 Then
        MsgBox "û������"
        InitData = -1
        Exit Function
    End If
    d.Arr = Range("A1").Resize(d.RowEnd, Pos.Cols).Value
    Set d.dic = CreateObject("Scripting.Dictionary") '�����ֵ���󣬺��ڰ󶨣�����Ҫ�����ã����ߡ����á������C:\WINDOWS\system32\scrrun.dll)
    
    Dim i As Long
    Dim str_key As String, strItem As String
    For i = Pos.RowStart To d.RowEnd
        str_key = VBA.CStr(d.Arr(i, Pos.TheName))
        '��Ŀ = Value + ��ע����עһ�������ݵĵ�λ��
        If d.dic.Exists(str_key) Then
            MsgBox VBA.Format(i, "��0�У��ظ�����Ŀ��")
            InitData = -1
            Exit Function
        Else
            strItem = Cells(i, Pos.Value).Text & VBA.CStr(d.Arr(i, Pos.Unit))
            If VBA.Len(strItem) Then
                d.dic(str_key) = strItem
            Else
                MsgBox "ֵΪ�գ����顣"
                Cells(i, Pos.Value).Select
                InitData = ErrRT
                Exit Function
            End If
        End If
    Next i
    
    '��word������Ѿ����˾Ͳ���Ҫ
    Set d.Doc = OpenDocument(d.DocFileName)
    If d.Doc Is Nothing Then
        InitData = ErrRT
        Exit Function
    End If
    
    InitData = SuccessRT
End Function

Sub AddComToWord()
    Dim d As DataStructDoc
    Dim i As Long
    
    If InitData(d) = -1 Then Exit Sub
    
    '�Ѿ��е���Ŀ�Ͳ���Ҫ��
    Dim str_key As String
    For i = 1 To d.Doc.Comments.Count
        str_key = d.Doc.Comments(i).Range.Text
        If d.dic.Exists(str_key) Then
            d.dic.Remove str_key
        End If
    Next
    
    Dim tmpKey
    tmpKey = d.dic.Keys()
    
    Dim p As Object '����
    For i = 0 To UBound(tmpKey)
        Set p = d.Doc.Paragraphs.Add
        p.Range.Text = tmpKey(i) & vbNewLine
        d.Doc.Comments.Add p.Range, VBA.CStr(tmpKey(i))
    Next
    
    MsgBox VBA.Format(i, "�����0����ע��")
End Sub
