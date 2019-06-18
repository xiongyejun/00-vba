Attribute VB_Name = "MUpdateWord"
Option Explicit

Enum WordField  
    wdFieldEmpty = -1 
End Enum

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

Sub UpdateWordByCom()  'ͨ������ʽ
    Dim i As Long, k As Long
    Dim f As Object, strKey As String
    Dim d As DataStructDoc
    
    '��ʼ������
    If InitData(d) = -1 Then Exit Sub
    
    If d.Doc.Fields.Count = 0 Then
        MsgBox "��ǰWordû����"
        GoTo A
    End If
    
    Dim strValue As String
    For i = 1 To d.Doc.Fields.Count
        Set f = d.Doc.Fields(i)
        strKey = GetKeyFromFieldCode(f.Code)
        
        If d.dic.Exists(strKey) Then
            strValue = d.dic(strKey)
            If f.Result.Text <> strValue Then
                f.Result.Text = strValue
            End If
        Else
            If MsgBox(VBA.Format(i, "��0����\[") & strValue & "]�������Ŀ��Excel��û�С�" & vbNewLine & "�Ƿ������", vbYesNo) = vbNo Then
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

Sub AddFieldToWord()
    Dim d As DataStructDoc
    Dim i As Long
    
    If InitData(d) = -1 Then Exit Sub
    
    '�Ѿ��е���Ŀ�Ͳ���Ҫ��
    Dim str_key As String
    For i = 1 To d.Doc.Fields.Count
        str_key = GetKeyFromFieldCode(d.Doc.Fields(i).Code)
        If d.dic.Exists(str_key) Then
            d.dic.Remove str_key
        End If
    Next
    
    Dim tmpKey, tmpItem
    tmpKey = d.dic.Keys()
    
    Dim p As Object '����
    Dim f As Object
    For i = 0 To UBound(tmpKey)
        Set p = d.Doc.Paragraphs.Add
        p.Range.Text = tmpKey(i) & vbNewLine
    Next
    
    Dim pCount As Long
    pCount = d.Doc.Paragraphs.Count
    For i = 0 To UBound(tmpKey)
        Set p = d.Doc.Paragraphs(pCount - i - 1)
        Set f = p.Range.Fields.Add(p.Range, WordField.wdFieldEmpty, VBA.CStr(tmpKey(i)))
        f.Result.Text = VBA.CStr(tmpKey(i))
        f.Result.Bold = False
    Next i
    
    MsgBox VBA.Format(i, "�����0����")
End Sub
'����code���ȡ�ؼ���key��excel�����Ŀ
Private Function GetKeyFromFieldCode(StrFieldCode As String) As String
    ' Key \* MERGEFORMAT
    GetKeyFromFieldCode = VBA.Mid$(StrFieldCode, 2, VBA.Len(StrFieldCode) - VBA.Len(" \* MERGEFORMAT ") - 1)
End Function
