Attribute VB_Name = "MOpenFile"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperationg As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Sub OpenFile()
    Dim FileName As String
    Dim rng As Range
    
    Set rng = Cells(ActiveCell.Row, "F")
    FileName = VBA.CStr(rng.Value)
    If VBA.Len(FileName) = 0 Then Exit Sub
    
    FileName = FindFile(FileName)
    If VBA.Len(FileName) Then
'        VBA.Shell "cmd.exe /c """ & fileName & """", vbNormalFocus
        If VBA.Right$(FileName, 4) = "xlsx" Or VBA.Right$(FileName, 4) = "xlsm" Or VBA.Right$(FileName, 3) = "xls" Then
            Workbooks.Open FileName, False
        Else
            Call ShellExecute(0&, vbNullString, FileName, vbNullString, vbNullString, vbNormalFocus)
        End If
    End If
End Sub

'�����ļ������ҵ�fullname����Ϊ��λǰ�������Ų�ȷ��
Function FindFile(FileName As String) As String
    Dim Path As String
    
    Path = ActiveWorkbook.Path
    
    Path = VBA.Left$(Path, VBA.InStrRev(Path, "\")) & "00-����\"
'    path = VBA.Left$(path, VBA.InStrRev(path, "\"))
    
    FindFile = Path & FileName
    If VBA.Len(VBA.Dir(FindFile)) Then
        Exit Function
    End If
    If VBA.Len(VBA.Dir(FindFile, vbDirectory)) Then
        Exit Function
    End If

    FindFile = ""
End Function
'�����ļ���·�����������㷢�ʼ�֮���
Sub CopyPath()
    Dim FileName As String
    Dim rng As Range
    
    Set rng = Cells(ActiveCell.Row, "F")
    FileName = VBA.CStr(rng.Value)
    If VBA.Len(FileName) = 0 Then Exit Sub
    
    FileName = FindFile(FileName)
    If VBA.Len(FileName) Then
        SetClipText FileName
    Else
        MsgBox "û���ҵ���Ӧ���ļ���"
    End If
End Sub
