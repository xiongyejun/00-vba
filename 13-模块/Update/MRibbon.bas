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
    InputFields Array("序号", "项目", "单位", "联系人", "电话", "文件名称", "时间", "备注")
End Sub

'------------------UpdateWord、PPT-----------------------------------
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
    InputFields Array("序号", "项目", "链接工作簿", "工作表", "单元格", "值", "单位", "PPTValue", "原表单位", "备注")
End Sub

'------------------UpdateWorkbbok-----------------------------------
Sub rbUpdateWorkbook(control As IRibbonControl)
    If CheckWorkbbokUpdateFields() Then Call UpdateWorkbook
End Sub
Sub rbOpenWorkbook(control As IRibbonControl)
    Dim rng As Range
    
    If CheckWorkbbokUpdateFields() Then
        Set rng = ActiveCell
        '如果选中的单元格是B列或者G列
        If rng.Column = Range("B1").Column Or rng.Column = Range("G1").Column Then
            If rng.Row > 1 Then
                If VBA.InStr(VBA.CStr(rng.Value), ":\") Then
                    Workbooks.Open VBA.CStr(rng.Value)
                Else
                    Workbooks.Open ActiveWorkbook.Path & "\" & VBA.CStr(rng.Value)
                End If
                
                If rng.Parent.Cells(rng.Row, "L").Value <> "AddName" And VBA.CStr(rng.Offset(0, 1).Value) <> "" Then ActiveWorkbook.Worksheets(VBA.CStr(rng.Offset(0, 1).Value)).Activate
            Else
                MsgBox "第1行应该是标题。"
            End If
        Else
            MsgBox "请选择B列或者G列"
        End If
    End If
End Sub
'添加类型的数据有效性
Sub rbAddValidation(control As IRibbonControl)
    Dim rng As Range
    
    If CheckWorkbbokUpdateFields() Then
        Set rng = Selection
        If rng.Columns.Count = 1 Then
            If rng.Column = Range("L1").Column Then
                rng.Validation.Delete
                rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="ColRelation,ColRelationAppend,Formula,Rng,dic,dicExists,AddName,AddNo"
            Else
                MsgBox "请选择L列。"
            End If
        Else
            MsgBox "不能选择多列数据。"
        End If
    End If
End Sub

Sub rbInputUpdateWorkbookFields(control As IRibbonControl)
    InputFields Array("序号", "源工作簿", "源工作表", "源起始行", "源定位列", "源单元格", "目标工作簿", "目标工作表", "目标起始行", "目标定位列", "目标单元格", "类型", "备注")
End Sub

'------------------通用项----------------------------------
Private Function CheckWordPPTUpdateFields() As Boolean
    CheckWordPPTUpdateFields = MFunc.CheckFields(Array("序号", "项目", "链接工作簿", "工作表", "单元格", "值", "单位", "PPTValue", "原表单位", "备注"))
End Function

Private Function CheckFileListFields() As Boolean
    CheckFileListFields = MFunc.CheckFields(Array("序号", "项目", "单位", "联系人", "电话", "文件名称", "时间", "备注"))
End Function

Private Function CheckWorkbbokUpdateFields() As Boolean
    CheckWorkbbokUpdateFields = MFunc.CheckFields(Array("序号", "源工作簿", "源工作表", "源起始行", "源定位列", "源单元格", "目标工作簿", "目标工作表", "目标起始行", "目标定位列", "目标单元格", "类型", "备注"))
End Function

'------------------其他-----------------------------------
Sub rbCloseMe(control As IRibbonControl)
    ThisWorkbook.Close False
End Sub


