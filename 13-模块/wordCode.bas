Sub SavePDF()
    Dim saveFileName As String
    Dim fd As FileDialog

    saveFileName = VBA.CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & ActiveDocument.Name & ".pdf"

    ActiveDocument.ExportAsFixedFormat saveFileName, wdExportFormatPDF
    MsgBox saveFileName
End Sub

Sub NoToText()
    Dim l As List
    
    If VBA.MsgBox("确定要将所有编号转为普通文本吗？", vbYesNo) = vbYes Then
        For Each l In ActiveDocument.Lists
            l.ConvertNumbersToText
        Next l
    End If
End Sub

Sub DeleteCom()
    Dim c As Comment
    Dim k As Long
    
    If VBA.MsgBox("确定要删除所有批注吗？", vbYesNo) = vbYes Then
        If VBA.MsgBox("删除前是否进行备份？", vbYesNo) = vbYes Then
            BackUp ActiveDocument
        End If
        
        k = 0
        For Each c In ActiveDocument.Comments
            c.Delete
            k = k + 1
        Next c
        
        MsgBox "已删除所有批注：" & k & "条。"
    End If
End Sub

Private Function BackUp(doc As Document)
    doc.Save
    Dim fileName As String, fileNameOld As String
    Dim k As Long
    
    fileNameOld = doc.FullName
    fileName = fileNameOld & ".bk" & VBA.CStr(k)
    Do Until VBA.Dir(fileName) = ""
        k = k + 1
        fileName = doc.FullName & ".bk" & VBA.CStr(k)
    Loop
    
    doc.SaveAs2 fileName
    doc.SaveAs2 fileNameOld

End Function
'切换全屏
Sub FullScreenQH()
    ActiveWindow.View.FullScreen = Not ActiveWindow.View.FullScreen
End Sub