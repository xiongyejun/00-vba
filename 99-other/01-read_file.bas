Attribute VB_Name = "read_file"
Option Explicit



Sub get_file_byte()
    Dim write_file_name As String
    Dim file_buffer() As Byte
    Dim file_name As String
    Dim i_mid As Long
    
    file_name = GetFileName
    If file_name = "" Then Exit Sub
    
    read_txt file_name, file_buffer
    
    Range("A:A").Clear
    i_mid = VBA.InStrRev(file_name, "\") + 1
    Range("A1").Value = VBA.Mid$(file_name, i_mid, VBA.Len(file_name) - i_mid + 1)
    
    byte_to_hex file_buffer
    
    MsgBox "OK"

End Sub

 Sub save_file_from_hex()
    Dim arr_data()
    Dim i_row As Long
    Dim fso As Object
    Dim save_file_name As String
    
    save_file_name = GetFolderPath
    If save_file_name = "" Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    i_row = Range("A" & Cells.Rows.Count).End(xlUp).Row
    arr_data = Range("A1:A" & i_row).Value
    
    save_file_name = save_file_name & "\" & arr_data(1, 1)
    
    fso.CreateTextFile save_file_name
    
    write_txt save_file_name, arr_data
    MsgBox "OK"
    
    Erase arr_data
    Set fso = Nothing
 
 End Sub
 
Function read_txt(txt_file_name As String, file_buffer() As Byte)
    Dim num_file As Integer
    
    num_file = FreeFile
    
    Open txt_file_name For Binary Access Read As #num_file
    ReDim file_buffer(LOF(num_file) - 1) As Byte
    Get #num_file, 1, file_buffer
    Close num_file
End Function
 
Function write_txt(txt_file_name As String, arr_data())
    Dim num_file As Integer
    Dim file_buffer() As Byte
    Dim i As Long, j As Long
    Dim str_hex As String
    
    num_file = FreeFile
    Open txt_file_name For Binary Access Write As #num_file
     
    For j = 2 To UBound(arr_data, 1)
        str_hex = arr_data(j, 1)
        
        ReDim file_buffer(VBA.Len(str_hex) / 2 - 1) As Byte
        
        For i = 1 To VBA.Len(str_hex) Step 2
            file_buffer((i - 1) / 2) = 0 + ("&H" & VBA.Mid$(str_hex, i, 2))
        Next i
        Put #num_file, , file_buffer
    Next j
    Close num_file
    
    Erase file_buffer
End Function
    
Function byte_to_hex(file_buffer() As Byte)
    Dim i As Long
    Dim str As String
    Dim k As Long
    
    k = 1
    For i = 0 To UBound(file_buffer)
        If i Mod 50 = 0 Then
            If i > 0 Then Cells(k, 1).Value = "'" & str
            str = ""
            k = k + 1
            str = VBA.Right$("00" & VBA.Hex(file_buffer(i)), 2)
        Else
            str = str & VBA.Right$("00" & VBA.Hex(file_buffer(i)), 2)
        End If
    Next i
    If str <> "" Then Cells(k, 1).Value = "'" & str
    
 End Function

Function GetFolderPath() As String
    Dim myFolder As Object
    Set myFolder = CreateObject("Shell.Application").Browseforfolder(0, "GetFolder", 0)
    If Not myFolder Is Nothing Then
'        GetFolderPath = myFolder.Items.item.path
        GetFolderPath = myFolder.Self.Path
        If Right(GetFolderPath, 1) <> "\" Then GetFolderPath = GetFolderPath & "\"
    Else
        GetFolderPath = ""
        MsgBox "请选择文件夹。"
    End If
    Set myFolder = Nothing
End Function

Function GetFileName(Optional strFilter As String = "*.*") As String
    With Application.FileDialog(msoFileDialogOpen)
'        .InitialFileName = ActiveWorkbook.path & "\" & strFilter
        If .Show = -1 Then                  ' -1代表确定，0代表取消
            GetFileName = .SelectedItems(1)
        Else
            GetFileName = ""
            MsgBox "请选择文件对象。"
        End If
    End With
End Function

