VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Implements MSForms.CommandButton

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public WithEvents tb As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1
Public WithEvents listBox As MSForms.listBox
Attribute listBox.VB_VarHelpID = -1
Public WithEvents cb As MSForms.CheckBox
Attribute cb.VB_VarHelpID = -1


Private Sub btn_Click()
    On Error GoTo Err1
    CallByName Me, btn.Tag, VbMethod
    Exit Sub
Err1:
    MsgBox Err.Description
End Sub


Private Sub cb_Click()
    Dim str As String
    str = cb.Parent.Parent.Controls(5).Text
    
    If cb.value Then
        cb.Parent.Parent.Controls(5).Text = str & " " & cb.Caption
    Else
        cb.Parent.Parent.Controls(5).Text = Replace(str, " " & cb.Caption, "")
    End If
End Sub

Private Sub listBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim Arr() As String
    
    Arr = adoGetFieldName(listBox.Tag, listBox.value)
    MsgBox Join(Arr, vbNewLine)
    Erase Arr
    
End Sub

Private Sub tb_Change()
    On Error GoTo Err1
    CallByName Me, tb.Tag, VbMethod
    Exit Sub
Err1:
    MsgBox Err.Description
End Sub

Sub tbSoure_Change()
    ThisWorkbook.Names.Add "数据源", tb.Text
    saveThisWorkbook
End Sub


 Sub btnCenterAcross_Click()    '跨列居中

    With Selection
        .UnMerge
        .Font.Bold = True
        .Font.Name = "宋体"
        .Font.Size = 18
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With
    Unload btn.Parent.Parent

End Sub

 Sub btnNumberFormatLocal_Click()       '设置单元格数字格式
    Selection.NumberFormatLocal = "#,##0.00;-#,##0.00;""-"";@"
    Unload btn.Parent.Parent
End Sub

 Sub btnStyle_Click()            '设置工作表样式
    With ActiveWorkbook.Styles("Normal")
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
    End With
    
    With ActiveWorkbook.Styles("Normal").Font
        .Name = "宋体"
        .Size = 12
        .Bold = False
        .Italic = False
        .Underline = xlUnderlineStyleNone
        .Strikethrough = False
        .ColorIndex = xlAutomatic
    End With
    
    With ActiveWorkbook.Styles("Normal")
        .IncludeNumber = True
        .IncludeFont = True
        .IncludeAlignment = True
        .IncludeBorder = True
        .IncludePatterns = True
        .IncludeProtection = True
    End With
    
    Cells.Style = "Normal"
    Unload btn.Parent.Parent
End Sub

 Sub btnAddButton_Click()       '工作表添加按钮
    Dim Shp As Button
    
    With ActiveCell
        Set Shp = ActiveSheet.Buttons.Add(.Left, .Top, .Width, .Height)
    End With
       
    Shp.Name = Application.InputBox("请输入按钮名称", Type:=2)
    Shp.Characters.Text = Shp.Name
    Unload btn.Parent.Parent

End Sub

 Sub btnJoinClipboard_Click()     '单元格区域连接起来
    Dim item, Arr1, Arr2(), k As Long
    Dim str As String
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "你选择的是" & TypeName(Selection) & "，我无法连接。"
        GoTo A
    End If
    
    If Selection.Cells.Count = 1 Then
        MsgBox "你只选择了一个单元格，不需要连接。"
        GoTo A
    End If
    
    Arr1 = Selection.value
    k = 0
    
    On Error GoTo A
    For Each item In Arr1
        ReDim Preserve Arr2(k)
        Arr2(k) = item
        k = k + 1
    Next item
    
    str = Application.InputBox("用什么字符连接呢？", Title:="输入字符", Default:="、", Type:=2)
    SetClipText Join(Arr2, str)
    
    Erase Arr1
    Erase Arr2
A:
    Unload btn.Parent.Parent

End Sub

 Sub btnChangeVersion_Click()               '更改为2003版本
    Dim path As String '路径
    Dim fileName As String '文件名称
    Dim fullName As String '2007文件

    path = ActiveWorkbook.path & "\"
    fileName = ActiveWorkbook.Name
    fullName = ActiveWorkbook.fullName

    If Right(fileName, 3) = "xls" Then
        MsgBox "文件是2003版本，无需更改", vbInformation
        GoTo A
    End If
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=path & Left(fileName, Len(fileName) - 4) & "xls", FileFormat:=xlNormal
        
    If MsgBox("是否删除2007文件？", vbQuestion + vbYesNo) = vbYes Then
        Kill fullName
    End If
    
    Application.DisplayAlerts = True
A:
    Unload btn.Parent.Parent
End Sub

 Sub btnSelectMerge_Click()
    Dim rng As Range
    Dim iCol As Long, iRow As Long, i As Long
    
    iCol = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
    iRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    For i = 1 To ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
        If IsNull(Columns(i).MergeCells) Then
            Exit For
        End If
    Next i
    If i = iCol + 1 Then GoTo A
    
    For Each rng In Range(Cells(1, i), Cells(iRow, i))
        Debug.Print rng.Address
        
        If rng.MergeCells = True Then
            rng.Select
            GoTo A
        End If
    Next rng
    
A:
    Set rng = Nothing
    Unload btn.Parent.Parent
End Sub

 Sub btnPasteValue_Click()           '数值粘贴
    Dim rng As Range, temp As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "请选择单元格区域。", vbInformation
        GoTo A
    End If
    
    If Selection.Cells.Count = 1 Then
        Selection.value = Selection.value
        GoTo A
    End If
    
    On Error Resume Next
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
    Set rng = rng.SpecialCells(xlCellTypeFormulas)
    rng.Select
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        Dim AreasRng As Range
        For Each AreasRng In rng.Areas
            AreasRng.value = AreasRng.value
        Next AreasRng
    End If
    
A:
    Set rng = Nothing
    Set temp = Nothing
    Set AreasRng = Nothing
    Unload btn.Parent.Parent
End Sub

 Sub btnUnProtectSht_Click()
    Dim sht As Worksheet
 
    For Each sht In Sheets
        With sht
            .Protect DrawingObjects:=True, Contents:=True, AllowFiltering:=True
            .Protect DrawingObjects:=False, Contents:=True, AllowFiltering:=True
            .Protect DrawingObjects:=True, Contents:=True, AllowFiltering:=True
            .Unprotect
       End With
    Next
    
    Unload btn.Parent.Parent
End Sub


'-----------------第二个Frame价格-----------------------------


'-----------------第三个Frame其他-----------------------------

 Sub btnQieHuanYinYong_Click()           '切换引用方式
    If Application.ReferenceStyle = xlR1C1 Then
        Application.ReferenceStyle = xlA1
    Else
        Application.ReferenceStyle = xlR1C1
    End If
    Unload btn.Parent.Parent
End Sub

 Sub btnBreakLink_Click()                '断开外部链接
    Dim Arr, i As Integer
    
    Arr = ActiveWorkbook.LinkSources(xlExcelLinks) '返回工作簿中链接的数组
    
    On Error GoTo A
    For i = 1 To UBound(Arr)
        ActiveWorkbook.BreakLink Name:=Arr(i), Type:=xlExcelLinks
    Next
A:
    Unload btn.Parent.Parent
End Sub
 Sub btnUnLoad_Click()                   'Unload btn.Parent.Parent
    Unload btn.Parent.Parent
End Sub

 Sub btnClose_Click()                    '关闭工作簿
    ThisWorkbook.Close False
End Sub
