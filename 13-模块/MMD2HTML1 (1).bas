Attribute VB_Name = "MMain"
Option Explicit

Sub vba_main()
    Dim path As String
    Dim md_path As String
    Dim ArrFile() As String
    Dim i As Long
    Dim html_name As String, md_name As String
    Dim k As Long
    Dim str_dir As String
    Dim str_next As String
    Dim str As String       '目录
    
    path = ThisWorkbook.path & "\"
    str_dir = GetFolderPath()
    k = MPublic.ScanDir(str_dir, ArrFile)
    
    str = GetDirStr(ArrFile)
    For i = 0 To k - 1
        md_name = ArrFile(i) & ".md"
        html_name = md_name & ".html"

        MdToHtml md_name, html_name, str

    Next i
    
    On Error GoTo err_handle
        
    Exit Sub
    
err_handle:
    MsgBox Err.Description
End Sub
'将md文件转换为html文件
'strDir    目录
Function MdToHtml(MdName As String, HtmlName As String, strDir As String)
    Dim num_file_md As Integer
    Dim num_file_html As Integer
    Dim str As String
    
    num_file_md = VBA.FreeFile
    Open MdName For Input As #num_file_md
    
    num_file_html = VBA.FreeFile
    Open HtmlName For Output As #num_file_html
    
    Do Until VBA.EOF(num_file_md)
        Line Input #num_file_md, str
        str = MdLineToHtml(str, num_file_md)
        Print #num_file_html, str
    Loop
    Print #num_file_html, strDir
    
    Close #num_file_md
    Close #num_file_html
End Function
'将md的文本行，转换为html格式
Function MdLineToHtml(StrMd As String, num_file_md As Integer) As String
    Dim i As Long
    Dim tmp1 As Long, tmp2 As Long
    Dim str_tmp As String
    
'    On Error GoTo err1
    
'    If VBA.InStr(StrMd, "person{""") Then Stop
'    For i = 1 To VBA.Len(StrMd)
'        Debug.Print Asc(VBA.Mid$(StrMd, i, 1))
'    Next
'
    If VBA.Left$(StrMd, 1) = "#" Then
        i = 2
        Do While VBA.Mid$(StrMd, i, 1) = "#"
            i = i + 1
        Loop
        i = i - 1
        MdLineToHtml = VBA.Format(i, "\<\h0\>") & VBA.Mid$(StrMd, i + 2) & VBA.Format(i, "\<\/\h0\>")
    
    ElseIf VBA.Left$(StrMd, 1) = "-" Then
        MdLineToHtml = "<li>" & VBA.Mid$(StrMd, 3) & "</li>"
        
    ElseIf VBA.InStr(StrMd, "![](../images") Then
    '图片
        StrMd = VBA.LTrim$(StrMd)
        MdLineToHtml = "<img src=""../" & VBA.Mid$(StrMd, 8, VBA.Len(StrMd) - 8) & """ />"
        
    ElseIf VBA.InStr(StrMd, "[目录]") Then
        MdLineToHtml = "<li><a href=""目录.html"">[目录]</a></li>"
        
    ElseIf VBA.InStr(StrMd, "](<") Then
    '上/下一节
        tmp1 = VBA.InStr(StrMd, "(<")
        tmp2 = VBA.InStr(StrMd, ">)")
        MdLineToHtml = "<li><a href=""" & VBA.Mid$(StrMd, tmp1 + 2, tmp2 - tmp1 - 2) & ".html"">"
        MdLineToHtml = MdLineToHtml & StrMd & "</a></li>"
        
    ElseIf VBA.Left$(StrMd, 3) = "```" Then
    'code
        StrMd = VBA.Replace(StrMd, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
        StrMd = "<span style=""color:#ee1b2e;""><ol><li>" & VBA.Mid$(StrMd, 4) & "<br/></li>" & vbNewLine
        Do
            StrMd = VBA.Replace(StrMd, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
            MdLineToHtml = MdLineToHtml & "<li>" & StrMd & "<br/></li>" & vbNewLine
            Line Input #num_file_md, StrMd
        Loop Until VBA.Left$(StrMd, 3) = "```"
        MdLineToHtml = MdLineToHtml & "</ol></span>"

    ElseIf VBA.Left$(StrMd, 1) = vbTab Then
    'TAB
        MdLineToHtml = "<p><span style=""color:#ee1b2e;"">" & VBA.Replace(StrMd, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;") & "</span></p>"
    
    Else
        MdLineToHtml = "<p>" & StrMd & "</p>"
    End If
    
err1:
    On Error GoTo 0
End Function


'获取目录字符
Function GetDirStr(ArrFile() As String) As String
    Dim arr() As String
    Dim k As Long, i As Long
    Dim num_file As Integer
    Dim str As String
    
    k = UBound(ArrFile)
    ReDim arr(k) As String
    For i = 0 To k
        arr(i) = VBA.Mid$(ArrFile(i), VBA.InStrRev(ArrFile(i), "\") + 1)   '去除文件名称
        '<li>ch0-03.md.html<a href="ch0-03.md.html">ch0-03.md.html</a></li>
        arr(i) = "<li>" & arr(i) & "<a href=""" & arr(i) & ".md.html""" & ">"
        '读取文件的第一行
        num_file = VBA.FreeFile
        Open ArrFile(i) & ".md" For Input As #num_file
        Line Input #num_file, str
        Close num_file
        
        arr(i) = arr(i) & str & "</a></li>"
    Next
    
    GetDirStr = VBA.Join(arr, vbNewLine)
End Function



