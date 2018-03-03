Attribute VB_Name = "MFunc"
Option Explicit

Function OpenDocument(FilePath As String) As Object
    Dim DocApp As Object 'New Word.Application
    Dim Doc As Object  'Word.Document
    
    On Error Resume Next
    Set DocApp = VBA.GetObject(, "Word.Application")
    If DocApp Is Nothing Then
        Set DocApp = VBA.CreateObject("Word.Application")
    End If
    On Error GoTo 0
    DocApp.Visible = True
    
    Dim fileName As String
    fileName = VBA.Right$(FilePath, VBA.Len(FilePath) - VBA.InStrRev(FilePath, "\"))
    
    '����򿪵�word����fileName�ģ���ʹ�����
    On Error Resume Next
    DocApp.Documents(fileName).Activate
    If Err.Number <> 0 Then
        Set Doc = DocApp.Documents.Open(FilePath)
    Else
        Set Doc = DocApp.Documents(fileName)
        If Doc.FullName <> FilePath Then
            MsgBox "����ͬ���ĵ����ˡ�"
            Set Doc = Nothing
        End If
    End If
    On Error GoTo 0
    
    Set OpenDocument = Doc
End Function

Function InitData(d As DataStruct) As Long
    shtLink.Activate
    'ѡ��word
    d.DocFileName = GetFileName()
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
    Set d.Dic = CreateObject("Scripting.Dictionary") '�����ֵ���󣬺��ڰ󶨣�����Ҫ�����ã����ߡ����á������C:\WINDOWS\system32\scrrun.dll)
    
    Dim i As Long
    Dim str_key As String
    For i = Pos.RowStart To d.RowEnd
        str_key = VBA.CStr(d.Arr(i, Pos.TheName))
        '��Ŀ = Value + ��ע����עһ�������ݵĵ�λ��
        If d.Dic.Exists(str_key) Then
            MsgBox VBA.Format(i, "��0�У��ظ�����Ŀ��")
            InitData = -1
            Exit Function
        Else
            d.Dic(str_key) = Cells(i, Pos.Value).Text & VBA.CStr(d.Arr(i, Pos.Cols))
        End If
    Next i
    
    '��word������Ѿ����˾Ͳ���Ҫ
    Set d.Doc = OpenDocument(d.DocFileName)
    If d.Doc Is Nothing Then
        InitData = -1
        Exit Function
    End If
    
    InitData = 1
End Function

Function GetFileName() As String
    With Application.FileDialog(msoFileDialogOpen)
        .InitialFileName = ActiveWorkbook.path & "\*.doc*"
        If .Show = -1 Then                  ' -1����ȷ����0����ȡ��
            GetFileName = .SelectedItems(1)
        Else
            GetFileName = ""
            MsgBox "��ѡ���ļ�����"
        End If
    End With
End Function
