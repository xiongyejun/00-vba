VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfFile 
   Caption         =   "UserForm2"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "usfFile.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "usfFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private frFolder As MSForms.Frame                            '文件夹框
Private WithEvents tbFolder As MSForms.TextBox
Attribute tbFolder.VB_VarHelpID = -1
Private WithEvents btnFolder As MSForms.CommandButton
Attribute btnFolder.VB_VarHelpID = -1

Private WithEvents frType As MSForms.Frame                  '文件类型框
Attribute frType.VB_VarHelpID = -1
Private cbSubfolder As MSForms.CheckBox                     '是否包含子文件夹
Private lbFileCount As MSForms.Label                        '文件数量

Private frBtn As MSForms.Frame                            '按钮框
Private WithEvents btnGetFile As MSForms.CommandButton    '获取文件
Attribute btnGetFile.VB_VarHelpID = -1
Private WithEvents btnPrint As MSForms.CommandButton      '输出文件列表
Attribute btnPrint.VB_VarHelpID = -1
Private WithEvents btnHide As MSForms.CommandButton       '隐藏窗体
Attribute btnHide.VB_VarHelpID = -1

Dim fileFullName() As String
Dim fileName() As String
Dim fileType() As String
Dim fileSize() As Long
Dim fileCount As Long
Dim typeD As Object

Private Sub btnFolder_Click()
    Dim str As String
    str = Func.GetFolderPath
    If str <> "" Then tbFolder.Text = str
End Sub

Private Sub btnGetFile_Click()
    If tbFolder.Text = "" Then Exit Sub
    
    Dim fso As Object
    Dim folder As Object
    Dim ctlArr()
    Dim iTop As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set typeD = CreateObject("Scripting.Dictionary")
    
    On Error GoTo Err
    
    Set folder = fso.GetFolder(tbFolder.Text)
    fileCount = 0
    
    lbFileCount.Caption = "正在查询…………"
    DoEvents
    Func.getFilesFromFolder fso, folder, fileFullName, fileName, fileType, fileSize, fileCount, typeD, cbSubfolder.value
    lbFileCount.Caption = "共有文件类型" & typeD.Count & "种，文件数：" & fileCount
    
    ctlArr = typeD.Keys
    iTop = Func.frameAddCheckBox(frType, ctlArr, 36)
    
    cbSubfolder.Top = iTop + 5
    iTop = iTop + cbSubfolder.Height
    lbFileCount.Top = iTop + 5
    iTop = iTop + lbFileCount.Height
    
    frBtn.Top = iTop + 5
    
    Me.Height = iTop + frBtn.Height + 35
'    Me.Left = (ActiveWindow.Width - Me.Width) / 2
'    Me.Top = (ActiveWindow.Height - Me.Height) / 2

A:
    Set fso = Nothing
    Set folder = Nothing
    Exit Sub
Err:
    MsgBox Err.Description
    GoTo A
    
End Sub

Private Sub btnHide_Click()
    Me.Hide
End Sub


Private Sub btnPrint_Click()
    Dim printArr()
    Dim i As Long, k As Long
    Dim ctl As Control
    Dim strType As String
    Dim rng As Range
    
    '0序号，1Path，2FileName，3Type,4Size,
    
    For Each ctl In frType.Controls
        If ctl.value Then
            k = k + typeD(ctl.Caption)
            strType = strType & "※" & ctl.Caption
        End If
    Next ctl
    
    If strType <> "" Then
        Me.Hide
        
        Func.getRngByInputBox rng
        If rng Is Nothing Then GoTo A
        
        ReDim printArr(k, 4)
        printArr(0, 0) = "No"
        printArr(0, 1) = "FullName"
        printArr(0, 2) = "FileName"
        printArr(0, 3) = "Type"
        printArr(0, 4) = "Size(K)"
        
        k = 1
        For i = 0 To fileCount - 1
            If InStr(strType, "※" & fileType(i)) > 0 Then
                printArr(k, 0) = k
                printArr(k, 1) = fileFullName(i)
                printArr(k, 2) = fileName(i)
                printArr(k, 3) = fileType(i)
                printArr(k, 4) = fileSize(i)
                k = k + 1
            End If
        Next i
        rng.Resize(k, 5).value = printArr
    Else
        MsgBox "你并没有选择任何文件类型。"
    End If
    
    
A:
    Set ctl = Nothing
    Erase printArr
End Sub

Private Sub frType_Click()
    Func.frameCheckBoxValue frType
End Sub

Private Sub frType_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Func.frameCheckBoxValue frType, False
End Sub

Private Sub tbFolder_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Shell "Explorer " & tbFolder.Text, 1
End Sub

Private Sub UserForm_Initialize()
    Const iWidth As Integer = 320
    Dim iTop As Integer
    Dim iLeft As Integer
    
    iTop = 5
    iLeft = 5
    Set frFolder = FrameAdd(Me, "Folder", iLeft, iTop, iWidth)
    Set tbFolder = tbAdd(frFolder, "", iLeft, iTop, iWidth - 35, 24)
    tbFolder.MultiLine = True
    Set btnFolder = btnAdd(frFolder, "浏览", tbFolder.Width, iTop, 30)
    frFolder.Height = tbFolder.Height + 20
    
    iTop = iTop + frFolder.Height + 5
    '文件类型
    Set frType = FrameAdd(Me, "文件类型", iLeft, iTop, iWidth)
    
    '是否包含子文件夹
    iTop = iTop + frType.Height + 5
    Set cbSubfolder = cbAdd(Me, "包含子文件夹", iLeft, iTop, iWidth)
    cbSubfolder.value = True
    
    '文件数量label
    iTop = iTop + cbSubfolder.Height + 5
    Set lbFileCount = labelAdd(Me, "", iLeft, iTop, Width)
    
    iTop = iTop + lbFileCount.Height + 5
    '按钮框
    Set frBtn = FrameAdd(Me, "", iLeft, iTop, iWidth)
    Set btnGetFile = btnAdd(frBtn, "获取文件", iLeft, 5)
    iLeft = iLeft + btnGetFile.Width
    Set btnPrint = btnAdd(frBtn, "输出", iLeft, 5)
    iLeft = iLeft + btnPrint.Width
    Set btnHide = btnAdd(frBtn, "隐藏", iLeft, 5)
    btnHide.Cancel = True
    
    frBtn.Height = btnPrint.Height + 10
    
    
    tbFolder.Text = ActiveWorkbook.path & "\"
    With Me
        .Width = iWidth + 15
        .Caption = "查询文件"
        .Height = iTop + frBtn.Height + 25
    End With

End Sub

