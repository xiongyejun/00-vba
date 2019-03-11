Attribute VB_Name = "MFunc"
Option Explicit

Function GetFileName(Optional strExt As String = "") As String
    With Application.FileDialog(msoFileDialogOpen)
        .InitialFileName = ActiveWorkbook.Path & "\*." & strExt & "*"
        If .Show = -1 Then                  ' -1����ȷ����0����ȡ��
            GetFileName = .SelectedItems(1)
        Else
            GetFileName = ""
            'MsgBox "��ѡ���ļ�����"
        End If
    End With
End Function

Function CheckFields(Fields As Variant) As Boolean
    Dim i As Long

    If VBA.IsArray(Fields) Then
        For i = 0 To UBound(Fields)
            If VBA.CStr(Cells(1, i + 1).Value) <> VBA.CStr(Fields(i)) Then
                MsgBox "������⣬A1��ʼ�ֱ��ǣ�" & vbNewLine & VBA.Join(Fields, "��")
                CheckFields = False
                Exit Function
            End If
        Next
    Else
        If VBA.CStr(Cells(1, 1).Value) <> VBA.CStr(Fields) Then
            MsgBox "������⣬A1=" & VBA.CStr(Fields)
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
        
        If MsgBox("ȷ����[" & Range("A1").Resize(1, i).Address(False, False) & "]������⣿" & vbNewLine & vbNewLine & VBA.Join(Fields, "��"), vbYesNo) = vbYes Then
            Range("A1").Resize(1, i).Value = Fields
        End If
    End If
End Function

Function SetClipText(str As String)
    Dim objData As Object 'New DataObject  '��Ҫ����"Microsoft Forms 2.0 Object Library"  FM20.DLL
    
    Set objData = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    With objData
        .SetText str       '�����ı�
        .PutInClipboard
      '  MsgBox "����ӵ������塣"
'        .GetFromClipboard               '��ȡ�ı�
'        MsgBox "��ǰ�������ڵ��ı��ǣ�" & .GetText
'        .Clear
'        .StartDrag
    End With
    Set objData = Nothing
    
End Function
