Attribute VB_Name = "MPublic"
Option Explicit

Function ScanDir(str_dir As String, ArrFile() As String) As Long
    Dim fso As Object
    Dim file As Object
    Dim folder As Object
    Dim k As Long
    Dim wk As Workbook
    
    On Error GoTo err_handle
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.Getfolder(str_dir)
    
    k = 0
    ReDim ArrFile(folder.Files.Count - 1) As String
    For Each file In folder.Files
        If VBA.InStr(file.Type, "MD") Then
            ArrFile(k) = VBA.Left$(file.path, VBA.Len(file.path) - 3) 'ȥ��.md��������
            k = k + 1
        End If
    Next file
    ReDim Preserve ArrFile(k - 1) As String
    
    InsertSort ArrFile, 0, k - 1
    
    ScanDir = k
    
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    Exit Function
    
err_handle:
    ScanDir = -1
End Function


Sub InsertSort(l() As String, Low As Long, High As Long)
    Dim i As Long, j As Long
    Dim ShaoBing  As String
     
    For i = Low + 1 To High
    
        If l(i) < l(i - 1) Then
            ShaoBing = l(i)             '�����ڱ�
                    
            j = i - 1
            Do While l(j) > ShaoBing
                l(j + 1) = l(j)
                j = j - 1
                If j = Low - 1 Then Exit Do
            Loop
            
            l(j + 1) = ShaoBing
        End If
    
    Next i
End Sub


Function GetFolderPath(Optional str_title As String = "��ѡ���ļ��С�") As String
    With Application.FileDialog(msoFileDialogFolderPicker)
'        .InitialFileName = ActiveWorkbook.path & "\"
        .Title = str_title
        
        If .Show = -1 Then                  ' -1����ȷ����0����ȡ��
            GetFolderPath = .SelectedItems(1)
        Else
            GetFolderPath = ""
            MsgBox "��ѡ���ļ��С�"
        End If
    End With
End Function

