Attribute VB_Name = "MUpdatePPT"
Option Explicit

Private Type DataStructPPT
    dic As Object '��¼����
    
    FileName As String
    pre As Object
End Type

Sub UpdatePPT()
    Dim d As DataStructPPT
    
    d.FileName = GetFileName("ppt")
    If d.FileName = "" Then Exit Sub
    
    Set d.pre = OpenPPT(d.FileName)
    If ReturnCode.ErrRT = ReadData(d) Then Exit Sub
    If ReturnCode.ErrRT = GetResult(d) Then Exit Sub
    
    MsgBox "OK"
End Sub

Private Function GetResult(d As DataStructPPT) As ReturnCode
    Dim s As Object ' Slide
    Dim shp As Object 'Shape
    Dim strKey As String
    
    For Each s In d.pre.Slides
        For Each shp In s.Shapes
            If shp.HasTextFrame Then
                strKey = shp.Name
'                Debug.Print shp.TextFrame.TextRange.Text
                If d.dic.Exists(strKey) Then
                    shp.TextFrame.TextRange.Text = d.dic(strKey)
                End If
            End If
        Next
    Next
    
    GetResult = SuccessRT
End Function

'��ȡ��Ŀ��Ӧ��PPTValue
'PPT�е�Shape����Ҫ����Ŀ����һ��
Private Function ReadData(d As DataStructPPT) As ReturnCode
    Set d.dic = VBA.CreateObject("Scripting.Dictionary")
    
    Dim Arr() As Variant
    Dim i_row As Long

    ActiveSheet.AutoFilterMode = False
    i_row = Cells(Cells.Rows.Count, Pos.TheName).End(xlUp).Row
    If i_row < Pos.RowStart Then
        MsgBox "û������"
        ReadData = ErrRT
        Exit Function
    End If
    Arr = Range("A1").Resize(i_row, Pos.Cols).Value
    
    Dim i As Long
    For i = Pos.RowStart To i_row
        d.dic(VBA.CStr(Arr(i, Pos.TheName))) = VBA.CStr(Arr(i, Pos.PPTValue))
    Next
    
    ReadData = SuccessRT
End Function

Private Function OpenPPT(FilePath As String) As Object
    Dim ppt As Object
    Dim pre As Object
    
    On Error Resume Next
    Set ppt = VBA.GetObject(, "Powerpoint.Application")
    If ppt Is Nothing Then
        Set ppt = VBA.CreateObject("Powerpoint.Application")
    End If
    On Error GoTo 0
    ppt.Visible = True
    
    Dim FileName As String
    FileName = VBA.Right$(FilePath, VBA.Len(FilePath) - VBA.InStrRev(FilePath, "\"))
    
    '����򿪵�ppt����fileName�ģ���ʹ�����
    On Error Resume Next
    Debug.Print ppt.Presentations(FileName).Name
    If Err.Number <> 0 Then
        Set pre = ppt.Presentations.Open(FilePath)
    Else
        Set pre = ppt.Presentations(FileName)
        If pre.FullName <> FilePath Then
            MsgBox "����ͬ���ĵ����ˡ�"
            Set ppt = Nothing
        End If
    End If
    On Error GoTo 0
    
    Set OpenPPT = pre
End Function

