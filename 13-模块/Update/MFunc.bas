Attribute VB_Name = "MFunc"
Option Explicit

Function GetFileName(Optional strExt As String = "") As String
    With Application.FileDialog(msoFileDialogOpen)
        .InitialFileName = ActiveWorkbook.Path & "\*." & strExt & "*"
        If .Show = -1 Then                  ' -1代表确定，0代表取消
            GetFileName = .SelectedItems(1)
        Else
            GetFileName = ""
            'MsgBox "请选择文件对象。"
        End If
    End With
End Function

Function CheckFields(Fields As Variant) As Boolean
    Dim i As Long

    If VBA.IsArray(Fields) Then
        For i = 0 To UBound(Fields)
            If VBA.CStr(Cells(1, i + 1).Value) <> VBA.CStr(Fields(i)) Then
                MsgBox "请检查标题，A1开始分别是：" & vbNewLine & VBA.Join(Fields, "、")
                CheckFields = False
                Exit Function
            End If
        Next
    Else
        If VBA.CStr(Cells(1, 1).Value) <> VBA.CStr(Fields) Then
            MsgBox "请检查标题，A1=" & VBA.CStr(Fields)
            CheckFields = False
            Exit Function
        End If
    End If
    
    CheckFields = True
End Function

Function InputFields(Fields As Variant) As Boolean
    Dim i As Long
    
    If VBA.IsArray(Fields) Then
        i = UBound(Fields) - LBound(Fields) + 1
        
        If MsgBox("确定在[" & Range("A1").Resize(1, i).Address(False, False) & "]输入标题？" & vbNewLine & vbNewLine & VBA.Join(Fields, "、"), vbYesNo) = vbYes Then
            Range("A1").Resize(1, i).Value = Fields
        End If
    End If
End Function

Function SetClipText(str As String)
    Dim objData As Object 'New DataObject  '需要引用"Microsoft Forms 2.0 Object Library"  FM20.DLL
    
    Set objData = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    With objData
        .SetText str       '设置文本
        .PutInClipboard
      '  MsgBox "已添加到剪贴板。"
'        .GetFromClipboard               '读取文本
'        MsgBox "当前剪贴板内的文本是：" & .GetText
'        .Clear
'        .StartDrag
    End With
    Set objData = Nothing
    
End Function
