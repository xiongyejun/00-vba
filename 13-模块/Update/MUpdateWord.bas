Attribute VB_Name = "MUpdateWord"
Option Explicit

Enum Pos
    RowStart = 2
    TheName = 2
    LinkWk = 3
    LinkSheet
    LinkRng
    Value
    Unit '数据的单位
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

Sub UpdateWordByCom()  '通过批注形式
    Dim i As Long, k As Long
    Dim Com, ComRngT As String, MyRange, iStart As Long, iEnd As Long
    Dim d As DataStructDoc
    
    '初始化数据
    If InitData(d) = -1 Then Exit Sub
    
    If d.Doc.Comments.Count = 0 Then
        MsgBox "当前Word没有批注。"
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
            If MsgBox(VBA.Format(i, "第0个批注\[") & ComRngT & "]：这个项目在Excel里没有。" & vbNewLine & "是否继续？", vbYesNo) = vbNo Then
                GoTo A
            End If
            
        End If

    Next i
        
    On Error Resume Next
    AppActivate ActiveWorkbook.Name
    On Error GoTo 0
    
    MsgBox "更新完成。"
      
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
    
    '如果打开的word存在fileName的，就使用这个
    On Error Resume Next
    DocApp.Documents(FileName).Activate
    If Err.Number <> 0 Then
        Set Doc = DocApp.Documents.Open(FilePath)
    Else
        Set Doc = DocApp.Documents(FileName)
        If Doc.FullName <> FilePath Then
            MsgBox "已有同名文档打开了。"
            Set Doc = Nothing
        End If
    End If
    On Error GoTo 0
    
    Set OpenDocument = Doc
End Function

Private Function InitData(d As DataStructDoc) As ReturnCode
    '选择word
    d.DocFileName = GetFileName("doc")
    If VBA.Len(d.DocFileName) = 0 Then
        InitData = -1
        Exit Function
    End If
    
    '读取数据
    ActiveSheet.AutoFilterMode = False
    d.RowEnd = Cells(Cells.Rows.Count, Pos.TheName).End(xlUp).Row
    If d.RowEnd < 2 Then
        MsgBox "没有数据"
        InitData = -1
        Exit Function
    End If
    d.Arr = Range("A1").Resize(d.RowEnd, Pos.Cols).Value
    Set d.dic = CreateObject("Scripting.Dictionary") '创建字典对象，后期绑定，不需要先引用（工具→引用→浏览→C:\WINDOWS\system32\scrrun.dll)
    
    Dim i As Long
    Dim str_key As String, strItem As String
    For i = Pos.RowStart To d.RowEnd
        str_key = VBA.CStr(d.Arr(i, Pos.TheName))
        '项目 = Value + 备注（备注一般是数据的单位）
        If d.dic.Exists(str_key) Then
            MsgBox VBA.Format(i, "第0行，重复的项目。")
            InitData = -1
            Exit Function
        Else
            strItem = Cells(i, Pos.Value).Text & VBA.CStr(d.Arr(i, Pos.Unit))
            If VBA.Len(strItem) Then
                d.dic(str_key) = strItem
            Else
                MsgBox "值为空，请检查。"
                Cells(i, Pos.Value).Select
                InitData = ErrRT
                Exit Function
            End If
        End If
    Next i
    
    '打开word，如果已经打开了就不需要
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
    
    '已经有的项目就不需要了
    Dim str_key As String
    For i = 1 To d.Doc.Comments.Count
        str_key = d.Doc.Comments(i).Range.Text
        If d.dic.Exists(str_key) Then
            d.dic.Remove str_key
        End If
    Next
    
    Dim tmpKey
    tmpKey = d.dic.Keys()
    
    Dim p As Object '段落
    For i = 0 To UBound(tmpKey)
        Set p = d.Doc.Paragraphs.Add
        p.Range.Text = tmpKey(i) & vbNewLine
        d.Doc.Comments.Add p.Range, VBA.CStr(tmpKey(i))
    Next
    
    MsgBox VBA.Format(i, "已添加0条批注。")
End Sub
