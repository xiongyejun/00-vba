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
    
    '如果打开的word存在fileName的，就使用这个
    On Error Resume Next
    DocApp.Documents(fileName).Activate
    If Err.Number <> 0 Then
        Set Doc = DocApp.Documents.Open(FilePath)
    Else
        Set Doc = DocApp.Documents(fileName)
        If Doc.FullName <> FilePath Then
            MsgBox "已有同名文档打开了。"
            Set Doc = Nothing
        End If
    End If
    On Error GoTo 0
    
    Set OpenDocument = Doc
End Function

Function InitData(d As DataStruct) As Long
    shtLink.Activate
    '选择word
    d.DocFileName = GetFileName()
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
    Set d.Dic = CreateObject("Scripting.Dictionary") '创建字典对象，后期绑定，不需要先引用（工具→引用→浏览→C:\WINDOWS\system32\scrrun.dll)
    
    Dim i As Long
    Dim str_key As String
    For i = Pos.RowStart To d.RowEnd
        str_key = VBA.CStr(d.Arr(i, Pos.TheName))
        '项目 = Value + 备注（备注一般是数据的单位）
        If d.Dic.Exists(str_key) Then
            MsgBox VBA.Format(i, "第0行，重复的项目。")
            InitData = -1
            Exit Function
        Else
            d.Dic(str_key) = Cells(i, Pos.Value).Text & VBA.CStr(d.Arr(i, Pos.Cols))
        End If
    Next i
    
    '打开word，如果已经打开了就不需要
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
        If .Show = -1 Then                  ' -1代表确定，0代表取消
            GetFileName = .SelectedItems(1)
        Else
            GetFileName = ""
            MsgBox "请选择文件对象。"
        End If
    End With
End Function
