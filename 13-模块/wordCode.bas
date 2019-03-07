Sub SavePDF()
    Dim saveFileName As String
    Dim fd As FileDialog

    saveFileName = VBA.CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & ActiveDocument.Name & ".pdf"

    ActiveDocument.ExportAsFixedFormat saveFileName, wdExportFormatPDF
    MsgBox saveFileName
End Sub

Sub NoToText()
    Dim l As List
    
    If VBA.MsgBox("ȷ��Ҫ�����б��תΪ��ͨ�ı���", vbYesNo) = vbYes Then
        For Each l In ActiveDocument.Lists
            l.ConvertNumbersToText
        Next l
    End If
End Sub

Sub DeleteCom()
    Dim c As Comment
    Dim k As Long
    
    If VBA.MsgBox("ȷ��Ҫɾ��������ע��", vbYesNo) = vbYes Then
        If VBA.MsgBox("ɾ��ǰ�Ƿ���б��ݣ�", vbYesNo) = vbYes Then
            BackUp ActiveDocument
        End If
        
        k = 0
        For Each c In ActiveDocument.Comments
            c.Delete
            k = k + 1
        Next c
        
        MsgBox "��ɾ��������ע��" & k & "����"
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
'�л�ȫ��
Sub FullScreenQH()
    ActiveWindow.View.FullScreen = Not ActiveWindow.View.FullScreen
End Sub