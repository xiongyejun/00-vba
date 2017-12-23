Attribute VB_Name = "MFunc"
Option Explicit

Function GetFileName(Optional str_tile As String = "", Optional str_filter As String = "") As String
    With Application.FileDialog(msoFileDialogOpen)
'        .InitialFileName = ActiveWorkbook.Path & "\"
        .Filters.Clear
        If VBA.Len(str_tile) > 0 Then .Title = str_tile
        
        If VBA.Len(str_filter) > 0 Then .Filters.Add VBA.Split(str_filter, "|")(0), VBA.Split(str_filter, "|")(1) 'CSV TXT|*.csv;*.txt
        
        If .Show = -1 Then                  ' -1����ȷ����0����ȡ��
            GetFileName = .SelectedItems(1)
        Else
            GetFileName = ""
'            MsgBox "��ѡ���ļ�����"
        End If
    End With
End Function

'�ж��ļ��Ƿ����
'FileName   �ļ�����
'IfMsg      �����ڵ�ʱ���Ƿ���ʾ��ʾ��Ϣ
Function FileExists(FileName As String, Optional IfMsg As Boolean = False) As Boolean
    If VBA.Len(FileName) = 0 Then
        FileExists = False
        Exit Function
    End If
    
    FileExists = VBA.Dir(FileName) <> ""
    
    If Not FileExists Then
        If IfMsg Then
            MsgBox "�����ڵ��ļ���" & vbNewLine & FileName
        End If
    End If
End Function
