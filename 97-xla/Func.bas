Attribute VB_Name = "Func"
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Function btnAdd(usf As Object, btnName As String, btnLeft As Integer, btnTop As Integer, Optional btnWidth As Integer = 72, Optional btnHeight As Integer = 24) As MSForms.CommandButton
    Dim btn As MSForms.CommandButton
    
    Set btn = usf.Controls.Add("Forms.CommandButton.1")
    With btn
        .Caption = btnName
        .Left = btnLeft
        .Width = btnWidth
        .Top = btnTop
        .Height = btnHeight
    End With
    
    Set btnAdd = btn
End Function

Function cbAdd(usf As Object, cbName As String, cbLeft As Integer, cbTop As Integer, Optional cbWidth As Integer = 108, Optional cbHeight As Integer = 18) As MSForms.CheckBox
    Dim cb As MSForms.CheckBox
    
    Set cb = usf.Controls.Add("Forms.CheckBox.1")
    With cb
        .Caption = cbName
        .Left = cbLeft
        .Width = cbWidth
        .Top = cbTop
        .Height = cbHeight
    End With
    
    Set cbAdd = cb
End Function

Function ComboBoxAdd(usf As Object, cbLeft As Integer, cbTop As Integer, Optional cbWidth As Integer = 72, Optional cbHeight As Integer = 18) As MSForms.ComboBox
    Dim cb As MSForms.ComboBox
    
    Set cb = usf.Controls.Add("Forms.ComboBox.1")
    With cb
        .Left = cbLeft
        .Width = cbWidth
        .Top = cbTop
        .Height = cbHeight
    End With
    
    Set ComboBoxAdd = cb
End Function

Function FrameAdd(usf As Object, frCaption As String, frLeft As Integer, frTop As Integer, Optional frWidth As Integer = 144, Optional frHeight As Integer = 216) As MSForms.Frame
    Dim fr As MSForms.Frame
    
    Set fr = usf.Controls.Add("Forms.Frame.1")
    With fr
        .Left = frLeft
        .Width = frWidth
        .Caption = frCaption
        .Top = frTop
        .Height = frHeight
    End With
    
    Set FrameAdd = fr
End Function

Function frameAddCheckBox(fr As MSForms.Frame, ctlArr(), Optional cbHeight As Integer = 18) As Long   '返回fr的底部Y坐标
    Dim i As Long
    Dim ctl As MSForms.CheckBox
    Const iWidth As Long = 80
    Const n As Long = 4
    
    fr.Controls.Clear
    For i = 0 To UBound(ctlArr)
        Set ctl = fr.Controls.Add("Forms.CheckBox.1")
        With ctl
            .Caption = ctlArr(i)
            .Top = cbHeight * (i \ n) + cbHeight / 4
            .Left = iWidth * (i Mod n)
            .Width = iWidth
            .Height = cbHeight
        End With
        
    Next i
    
    With fr
        .Width = n * iWidth
        .Height = cbHeight * (i \ n + 1) + cbHeight / 2
        If .Height > 300 Then
            .ScrollHeight = .Height
            .Height = 300 '太大了超过了屏幕看不到
            .ScrollBars = fmScrollBarsVertical
        Else
            .ScrollBars = fmScrollBarsNone
        End If
        frameAddCheckBox = .Top + .Height
    End With
    
End Function

Function frameCheckBoxValue(fr As MSForms.Frame, Optional value As Boolean = True)
    Dim ctl As Control
    For Each ctl In fr.Controls
        ctl.value = value
    Next ctl
    Set ctl = Nothing
End Function


Function labelAdd(usf As Object, lbName As String, lbLeft As Integer, lbTop As Integer, Optional lbWidth As Integer = 72, Optional lbHeight As Integer = 18) As MSForms.Label
    Dim lb As MSForms.Label
    
    Set lb = usf.Controls.Add("Forms.Label.1")
    With lb
        .Caption = lbName
        .Left = lbLeft
        .Width = lbWidth
        .Top = lbTop
        .Height = lbHeight
    End With
    
    Set labelAdd = lb
End Function

Function listBoxAdd(usf As Object, lbLeft As Integer, lbTop As Integer, Optional lbWidth As Integer = 72, Optional lbHeight As Integer = 72) As MSForms.listBox
    Dim lb As MSForms.listBox
    
    Set lb = usf.Controls.Add("Forms.ListBox.1")
    With lb
        .Left = lbLeft
        .Width = lbWidth
        .Top = lbTop
        .Height = lbHeight
    End With
    
    Set listBoxAdd = lb
End Function

Function tbAdd(usf As Object, tbName As String, tbLeft As Integer, tbTop As Integer, Optional tbWidth As Integer = 72, Optional tbHeight As Integer = 18) As MSForms.TextBox
    Dim tb As MSForms.TextBox
    
    Set tb = usf.Controls.Add("Forms.TextBox.1")
    With tb
        .Text = tbName
        .Left = tbLeft
        .Width = tbWidth
        .Top = tbTop
        .Height = tbHeight
    End With
    
    Set tbAdd = tb
End Function

Function getRngByInputBox(rng As Range, Optional strPrompt As String = "选择输出单元格。")
    On Error Resume Next
    Set rng = Application.InputBox(strPrompt, Default:=ActiveCell.Address, Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "请选择单元格区域。"
    Else
        Set rng = rng.Range("a1")
    End If

End Function

Function getSheetNameByAdo(fileName As String, shtNameArr() As String) '0表示出错，k表示工作表数量
    Dim AdoConn As Object ' New ADODB.Connection
    Dim AdoRst As Object ' ADODB.Recordset
    Dim StrConn As String
    Dim strSql As String
    Dim k As Long
    
    On Error GoTo Err1:
    
    StrConn = ExcelData(fileName)
    Set AdoConn = CreateObject("ADODB.Connection")
    AdoConn.Open StrConn
    Set AdoRst = CreateObject("ADODB.Recordset")
    Set AdoRst = AdoConn.OpenSchema(20) 'adSchemaTables
    
    k = 0
    Do Until AdoRst.EOF
        If AdoRst!Table_type = "TABLE" And AdoRst!TABLE_NAME Like "*$" Then
            k = k + 1
            ReDim Preserve shtNameArr(1 To k) As String
            shtNameArr(k) = AdoRst!TABLE_NAME
        End If
        AdoRst.MoveNext
    Loop
    
    getSheetNameByAdo = k
A:
    Set AdoConn = Nothing
    Set AdoRst = Nothing
    Exit Function
Err1:
    MsgBox Err.Description
    getSheetNameByAdo = 0
    GoTo A
End Function

Function getFieldNameByAdo(fileName As String, tableName As String, fieldNameArr() As String) '0表示出错，k表示字段数量
    Dim AdoConn As Object ' New ADODB.Connection
    Dim AdoRst As Object ' ADODB.Recordset
    Dim StrConn As String
    Dim strSql As String
    Dim k As Long
    
    On Error GoTo Err1:
    
    StrConn = ExcelData(fileName)
    Set AdoConn = CreateObject("ADODB.Connection")
    AdoConn.Open StrConn
    Set AdoRst = CreateObject("ADODB.Recordset")
    Set AdoRst = AdoConn.OpenSchema(4) 'adSchemaColumns
    
    k = 0
    Do Until AdoRst.EOF
        If AdoRst!TABLE_NAME = tableName Then
            k = k + 1
            ReDim Preserve fieldNameArr(1 To k) As String
            fieldNameArr(k) = AdoRst!COLUMN_NAME
        End If
        AdoRst.MoveNext
    Loop
    
    adoGetFieldName = k
A:
    Set AdoConn = Nothing
    Set AdoRst = Nothing
    Exit Function
Err1:
    MsgBox Err.Description
    adoGetFieldName = 0
    GoTo A
End Function

Function getShtNameFromWorkbook(wk As Workbook, Arr() As String, Optional ifGetRow1RangeAddress As Boolean = False)
    Dim i As Long
    
    ReDim Arr(wk.Worksheets.Count)
    For i = 1 To wk.Worksheets.Count
        With wk.Worksheets(i)
            Arr(i) = .Name
            If ifGetRow1RangeAddress Then
                Arr(i) = Arr(i) & "$A:"
                Arr(i) = Arr(i) & Split(.Range("IV1").End(xlToLeft).Address, "$")(1)
            End If
        End With
    Next i
End Function
Function SetClipText(str As String)
    Dim objData As New DataObject  '需要引用"Microsoft Forms 2.0 Object Library"  FM20.DLL
    
    With objData
        .SetText str       '设置文本
        .PutInClipboard
        MsgBox "已添加到剪贴板。"
'        .GetFromClipboard               '读取文本
'        MsgBox "当前剪贴板内的文本是：" & .GetText
'        .Clear
'        .StartDrag
    End With
    Set objData = Nothing
    
End Function

Function setFormPosLeft(formWidth As Long) As Long
    Dim x As Long
    x = GetSystemMetrics(SM_CXSCREEN)
    
    setFormPosLeft = (x - formWidth) / 2
    
End Function

Function setFormPosRight(formHeight As Long) As Long
    Dim y As Long
    y = GetSystemMetrics(SM_CYSCREEN)
    setFormPosRight = (y - formHeight) / 2
    
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

Function GetFolderPath() As String
    Dim myFolder As Object
    Set myFolder = CreateObject("Shell.Application").Browseforfolder(0, "GetFolder", 0)
    If Not myFolder Is Nothing Then
'        GetFolderPath = myFolder.Items.item.path
        GetFolderPath = myFolder.Self.path
        If Right(GetFolderPath, 1) <> "\" Then GetFolderPath = GetFolderPath & "\"
    Else
        GetFolderPath = ""
        MsgBox "请选择文件夹。"
    End If
    Set myFolder = Nothing
End Function

Function getWorkbook(wk As Workbook, strWk As String) As Long '获取一个工作簿对象，0表示出错，1表示是重新打开的，2表示工作簿是已经打开的
    Dim i As Long, iWk As Long
    
    getWorkbook = 2
    
    On Error GoTo ErrHandle
    iWk = Workbooks.Count
    For i = 1 To iWk
        If Workbooks(i).fullName = strWk Then
            Set wk = Workbooks(i)
            Exit For
        End If
    Next i
    
    If i = iWk + 1 Then             '说明工作簿没有打开
        Set wk = Workbooks.Open(strWk, UpdateLinks:=False)
        getWorkbook = 1
    End If
    
    Exit Function
ErrHandle:
    MsgBox Err.Description
    getWorkbook = 0
End Function

Function saveThisWorkbook()
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
End Function

Function ExcelData(fileName As String) As String '读取Excel数据库
    If Val(Application.Version) > 11 Then
        ExcelData = "OLEDB;Provider =Microsoft.ACE.OLEDB.12.0;Data Source=" _
                    & fileName & ";Extended Properties=""Excel 12.0;HDR=YES"";"
    Else
        ExcelData = "OLEDB;Provider =Microsoft.Jet.OLEDB.4.0;Data Source=" _
                    & fileName & ";Extended Properties=""Excel 8.0;HDR=YES"";"
    End If
    
    
End Function
Function AccessData(fileName As String) As String '读取Aceess数据库
   AccessData = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fileName
End Function
'Microsoft ActiveX Data Objects 2.8 Library
Function CreateAdo(SqlStr As String, rng As Range, fileName As String) As Long  '0表示出错，1表示正确
    Dim AdoConn As Object, rst As Object
    Dim i As Long
    
    On Error GoTo Err
    Set AdoConn = CreateObject("ADODB.Connection")
    Set rst = CreateObject("ADODB.Recordset")
    
    AdoConn.Open ExcelData(fileName)
    rst.Open SqlStr, AdoConn
    
    For i = 0 To rst.Fields.Count - 1
        rng.Offset(0, i).value = rst.Fields(i).Name
    Next i

    rng.Offset(1, 0).CopyFromRecordset rst 'AdoConn.Execute(SqlStr)
    CreateAdo = 1
    rst.Close
    
A:
    AdoConn.Close
    Set rst = Nothing
    Set AdoConn = Nothing
    Exit Function
    
Err:
    MsgBox Err.Description
    CreateAdo = 0
    GoTo A
End Function

Function getFilesFromFolder(fso As Object, folder As Object, _
                            fileFullName() As String, fileName() As String, fileType() As String, _
                            fileSize() As Long, fileCount As Long, typeD As Object, Optional ifSubfolder As Boolean = True)
                            
    Dim file As Object
    Dim subFolders As Object, subFolder As Object
    
    On Error Resume Next            '有些没有权限的
    For Each file In folder.Files
        If Err.Number = 0 Then
            ReDim Preserve fileFullName(fileCount) As String
            ReDim Preserve fileName(fileCount) As String
            ReDim Preserve fileType(fileCount) As String
            ReDim Preserve fileSize(fileCount) As Long
            
            With file
                fileFullName(fileCount) = .path
                fileName(fileCount) = .Name
                fileType(fileCount) = .Type
                fileSize(fileCount) = .Size / 1024
            End With
            typeD(fileType(fileCount)) = typeD(fileType(fileCount)) + 1
            
            fileCount = fileCount + 1
        End If
        Err.Clear
    Next file
    
    If ifSubfolder Then
        Set subFolders = folder.subFolders
        If subFolders.Count > 0 Then                '循环文件夹
            For Each subFolder In subFolders
                getFilesFromFolder fso, subFolder, fileFullName, fileName, fileType, fileSize, fileCount, typeD, ifSubfolder
            Next subFolder
        End If
    End If
    On Error GoTo 0
    
    Set file = Nothing
    Set subFolders = Nothing
    Set subFolder = Nothing
End Function
