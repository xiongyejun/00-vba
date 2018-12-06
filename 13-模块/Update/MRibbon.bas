Attribute VB_Name = "MRibbon"
Option Explicit

Sub RibbonUI_onLoad(Ribbon As IRibbonUI)
    On Error Resume Next
    Ribbon.ActivateTab "TabIDUpdate"
End Sub

Sub rbOpenFile(control As IRibbonControl)
    If CheckFileListFields() Then Call OpenFile
End Sub

Sub rbCopyPath(control As IRibbonControl)
    If CheckFileListFields() Then Call CopyPath
End Sub

Sub rbInputFileListFields(control As IRibbonControl)
    InputFields Array("���", "��Ŀ", "��λ", "��ϵ��", "�绰", "�ļ�����", "ʱ��", "��ע")
End Sub

'------------------UpdateWord��PPT-----------------------------------
Sub rbUpdateExcel(control As IRibbonControl)
    If CheckWordPPTUpdateFields() Then Call UpdateExcel
End Sub

Sub rbUpdateWord(control As IRibbonControl)
    If CheckWordPPTUpdateFields() Then Call UpdateWordByCom
End Sub

Sub rbAddComToWord(control As IRibbonControl)
    If CheckWordPPTUpdateFields() Then Call AddComToWord
End Sub

Sub rbUpdatePPT(control As IRibbonControl)
    If CheckWordPPTUpdateFields() Then Call UpdatePPT
End Sub

Sub rbInputUpdateFields(control As IRibbonControl)
    InputFields Array("���", "��Ŀ", "���ӹ�����", "������", "��Ԫ��", "ֵ", "��λ", "PPTValue", "ԭ��λ", "��ע")
End Sub

'------------------UpdateWorkbbok-----------------------------------
Sub rbUpdateWorkbook(control As IRibbonControl)
    If CheckWorkbbokUpdateFields() Then Call UpdateWorkbook
End Sub
Sub rbOpenWorkbook(control As IRibbonControl)
    Dim rng As Range
    
    If CheckWorkbbokUpdateFields() Then
        Set rng = ActiveCell
        '���ѡ�еĵ�Ԫ����B�л���G��
        If rng.Column = Range("B1").Column Or rng.Column = Range("G1").Column Then
            If rng.Row > 1 Then
                If VBA.InStr(VBA.CStr(rng.Value), ":\") Then
                    Workbooks.Open VBA.CStr(rng.Value)
                Else
                    Workbooks.Open ActiveWorkbook.Path & "\" & VBA.CStr(rng.Value)
                End If
                
                If rng.Parent.Cells(rng.Row, "L").Value <> "AddName" And VBA.CStr(rng.Offset(0, 1).Value) <> "" Then ActiveWorkbook.Worksheets(VBA.CStr(rng.Offset(0, 1).Value)).Activate
            Else
                MsgBox "��1��Ӧ���Ǳ��⡣"
            End If
        Else
            MsgBox "��ѡ��B�л���G��"
        End If
    End If
End Sub
'������͵�������Ч��
Sub rbAddValidation(control As IRibbonControl)
    Dim rng As Range
    
    If CheckWorkbbokUpdateFields() Then
        Set rng = Selection
        If rng.Columns.Count = 1 Then
            If rng.Column = Range("L1").Column Then
                rng.Validation.Delete
                rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="ColRelation,ColRelationAppend,Formula,Rng,dic,dicExists,AddName,AddNo"
            Else
                MsgBox "��ѡ��L�С�"
            End If
        Else
            MsgBox "����ѡ��������ݡ�"
        End If
    End If
End Sub

Sub rbInputUpdateWorkbookFields(control As IRibbonControl)
    InputFields Array("���", "Դ������", "Դ������", "Դ��ʼ��", "Դ��λ��", "Դ��Ԫ��", "Ŀ�깤����", "Ŀ�깤����", "Ŀ����ʼ��", "Ŀ�궨λ��", "Ŀ�굥Ԫ��", "����", "��ע")
End Sub

'------------------ͨ����----------------------------------
Private Function CheckWordPPTUpdateFields() As Boolean
    CheckWordPPTUpdateFields = MFunc.CheckFields(Array("���", "��Ŀ", "���ӹ�����", "������", "��Ԫ��", "ֵ", "��λ", "PPTValue", "ԭ��λ", "��ע"))
End Function

Private Function CheckFileListFields() As Boolean
    CheckFileListFields = MFunc.CheckFields(Array("���", "��Ŀ", "��λ", "��ϵ��", "�绰", "�ļ�����", "ʱ��", "��ע"))
End Function

Private Function CheckWorkbbokUpdateFields() As Boolean
    CheckWorkbbokUpdateFields = MFunc.CheckFields(Array("���", "Դ������", "Դ������", "Դ��ʼ��", "Դ��λ��", "Դ��Ԫ��", "Ŀ�깤����", "Ŀ�깤����", "Ŀ����ʼ��", "Ŀ�궨λ��", "Ŀ�굥Ԫ��", "����", "��ע"))
End Function

'------------------����-----------------------------------
Sub rbCloseMe(control As IRibbonControl)
    ThisWorkbook.Close False
End Sub


