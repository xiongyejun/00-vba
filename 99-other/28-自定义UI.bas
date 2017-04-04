Attribute VB_Name = "模块1"
Option Explicit

Sub vba_main()
    Dim i As Long
    Dim i_row As Long
    Dim Arr()
    Dim action_path As String
    Dim str_input As String
    Dim str As String
    
    ActiveSheet.AutoFilterMode = False
    i_row = Range("A" & Cells.Rows.Count).End(xlUp).Row
    Arr = Range("A1：D" & i_row).Value
   
    str = "<mso:cmd app=""Excel"" dt=""1"" />" & vbNewLine

    str = str & vbNewLine & "<mso:customUI xmlns:x1=""http://schemas.microsoft.com/office/2009/07/customui/macro"" xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">"
    str = str & vbNewLine & "  <mso:ribbon>"
    str = str & vbNewLine & "  <mso:tabs>"
   
    For i = 2 To i_row
        If VBA.Left(Arr(i, 1), 2) = "</" Then
            str_input = VBA.Space$(4) & Arr(i, 1)
            str = str & vbNewLine & str_input
            str = str & vbNewLine
        ElseIf VBA.Left(Arr(i, 1), 10) = "<mso:group" Then
            action_path = Arr(i, 4)
            str_input = VBA.Space$(4) & Arr(i, 1)
           str = str & vbNewLine & str_input
        ElseIf VBA.Left(Arr(i, 1), 1) = "<" Then
            str_input = VBA.Space$(2) & Arr(i, 1)
            str = str & vbNewLine & str_input
            str = str & vbNewLine
        Else
            str_input = VBA.Space$(5) & "<mso:button idQ=""x1:" & Arr(i, 1) & i & """ label=""" & _
                            Arr(i, 2) & """ imageMso=""" & Arr(i, 3) & _
                            """ onAction=""" & action_path & Arr(i, 4) & """ visible=""true""/>"
            str = str & vbNewLine & str_input
        End If
    Next i
   
    str = str & vbNewLine & "   </mso:tabs>"
    str = str & vbNewLine & "  </mso:ribbon>"
    str = str & vbNewLine & "</mso:customUI>"
'    Debug.Print str
    
    WriteUTF8 get_my_doc() & "\Excel 自定义.exportedUI", str

'    VBA.Shell "C:\WINDOWS\NOTEPAD.EXE " & get_my_doc() & "\Excel 自定义.exportedUI", vbNormalFocus
End Sub

Sub WriteUTF8(strPath As String, str As String)
    Dim WriteStream As Object
    Set WriteStream = CreateObject("ADODB.Stream")
    With WriteStream
        .Type = 2               'adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText str
        .SaveToFile strPath, 2  'adSaveCreateOverWrite
        .Flush
        .Close
        
    End With
    Set WriteStream = Nothing
    
End Sub

Function get_my_doc() As String
    Dim wsh As Object
    Dim str As String
    
    Set wsh = CreateObject("WScript.Shell")
    str = wsh.SpecialFolders("Mydocuments")
    get_my_doc = str
    
    Set wsh = Nothing
End Function
