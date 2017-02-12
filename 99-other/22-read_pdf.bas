Attribute VB_Name = "模块1"
Sub cmd_imp_Click()

    Dim OS_FSO As Object
    Set OS_FSO = CreateObject("Scripting.filesystemobject")
    Dim Dlg_File As FileDialog
    Dim PDF_Path As String
    
    Set Dlg_File = Application.FileDialog(msoFileDialogFilePicker)
    
    With Dlg_File
        .Filters.Add "PDF文件", "*.pdf"
        If .Show = -1 Then
            PDF_Path = .SelectedItems(1)
        End If
    End With
    
    If OS_FSO.fileexists(PDF_Path) = False Then
        MsgBox "PDF文件没有找到"
        Set OS_FSO = Nothing
        Exit Sub
    End If
    
    ReadPDFToExcel PDF_Path

End Sub

Function ReadPDFToExcel(PDF_File As String)
    Dim AC_PD As Acrobat.AcroPDDoc
    Dim AC_Hi As Acrobat.AcroHiliteList
    Dim AC_PG As Acrobat.AcroPDPage
    Dim AC_PGTxt As Acrobat.AcroPDTextSelect
    
    Dim rng As Range
    
    Dim Ct_Page As Long
    Dim i As Long, j As Long, k As Long
    Dim T_Str As String
    Dim Hld_Txt As Variant
    
    Application.ScreenUpdating = False
    Cells.Clear
    
    Set AC_PD = New Acrobat.AcroPDDoc
    Set AC_Hi = New Acrobat.AcroHiliteList 'Hilite醒目 List
    
'    adds the specified highlight to the current highlight list.
'    添加指定的突出当前突出显示列表。
    AC_Hi.Add 0, 32767
    
    With AC_PD
       .Open PDF_File
        Ct_Page = .GetNumPages
        
        If Ct_Page = -1 Then
            MsgBox "请确认PDF文件 '" & PDF_File & "'"
            .Close
            GoTo h_end
        End If
    
        For i = 1 To Ct_Page
            T_Str = ""
            'acquires the specified page.
            '获得指定的页面。
            Set AC_PG = .AcquirePage(i - 1)
            'Creates a text selection on a single page.
            '创建一个文本选择在一个页面。
            Set AC_PGTxt = AC_PG.CreateWordHilite(AC_Hi)
                    
            If Not AC_PGTxt Is Nothing Then
                With AC_PGTxt
                    'Gets the number of text elements in a text selection.
                    For j = 0 To .GetNumText - 1
                        ' Gets the text from the specified element of a text selection.
                        T_Str = T_Str & .GetText(j)
                    Next j
                End With
            End If
            
            With WS_PDF
                Set rng = Range("A" & Cells.Rows.Count).End(xlUp).Offset(1, 0)
                rng.Offset(1, 0).Value = VBA.Format(i, "第0页")
                
                If T_Str <> "" Then
                    Hld_Txt = VBA.Split(T_Str, vbNewLine)
                    rng.Offset(2, 0).Resize(UBound(Hld_Txt) + 1, 1).Value = Application.WorksheetFunction.Transpose(Hld_Txt)
                End If
            End With
        Next i
        
        .Close
        
    End With
                
    Application.ScreenUpdating = True
    
    MsgBox "完成"
                
h_end:
        
    Set WS_PDF = Nothing
    Set AC_PGTxt = Nothing
    Set AC_PG = Nothing
    Set AC_Hi = Nothing
    Set AC_PD = Nothing
End Function

