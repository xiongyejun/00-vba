Attribute VB_Name = "MFunc"
Option Explicit

Function GetFileName(Optional str_tile As String = "", Optional str_filter As String = "") As String
    With Application.FileDialog(msoFileDialogOpen)
'        .InitialFileName = ActiveWorkbook.Path & "\"
        .Filters.Clear
        If VBA.Len(str_tile) > 0 Then .Title = str_tile
        
        If VBA.Len(str_filter) > 0 Then .Filters.Add VBA.Split(str_filter, "|")(0), VBA.Split(str_filter, "|")(1) 'CSV TXT|*.csv;*.txt
        
        If .Show = -1 Then                  ' -1代表确定，0代表取消
            GetFileName = .SelectedItems(1)
        Else
            GetFileName = ""
'            MsgBox "请选择文件对象。"
        End If
    End With
End Function

'判断文件是否存在
'FileName   文件名称
'IfMsg      不存在的时候是否显示提示消息
Function FileExists(FileName As String, Optional IfMsg As Boolean = False) As Boolean
    If VBA.Len(FileName) = 0 Then
        FileExists = False
        Exit Function
    End If
    
    FileExists = VBA.Dir(FileName) <> ""
    
    If Not FileExists Then
        If IfMsg Then
            MsgBox "不存在的文件：" & vbNewLine & FileName
        End If
    End If
End Function
