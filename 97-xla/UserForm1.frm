VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mc() As New MyControl

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Erase mc
End Sub

Private Sub UserForm_Initialize()                '常用命令
    Dim btnArr(1 To 50, 1 To 3) As String
    Const iLeft As Integer = 2, iWidth As Integer = 90, iHeight As Integer = 24
    Dim iBtnTop As Integer, i As Integer, iBtn As Integer
    Dim btn As MSForms.CommandButton
    Dim fmFormat As Frame, fmJiaGe As Frame, fmOther As Frame
    Dim fr As Control, iFrTop As Integer
    
    With Me.Controls
        Set fmFormat = .Add("Forms.Frame.1"): fmFormat.Caption = "格式"
        Set fmJiaGe = .Add("Forms.Frame.1"): fmJiaGe.Caption = "Two"
        Set fmOther = .Add("Forms.Frame.1"): fmOther.Caption = "其他"
    End With
    
    iBtn = 1
    '1-btn.Caption;2-btn.Click;3-btn.Parent
    'fmFormat
    btnArr(iBtn, 1) = "跨列居中": btnArr(iBtn, 2) = "btnCenterAcross_Click": btnArr(iBtn, 3) = "fmFormat": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "NumberFormatLocal": btnArr(iBtn, 2) = "btnNumberFormatLocal_Click": btnArr(iBtn, 3) = "fmFormat": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "Style常规": btnArr(iBtn, 2) = "btnStyle_Click": btnArr(iBtn, 3) = "fmFormat": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "窗体按钮Add": btnArr(iBtn, 2) = "btnAddButton_Click": btnArr(iBtn, 3) = "fmFormat": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "JoinClipboard": btnArr(iBtn, 2) = "btnJoinClipboard_Click": btnArr(iBtn, 3) = "fmFormat": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "SaveAs2003": btnArr(iBtn, 2) = "btnChangeVersion_Click": btnArr(iBtn, 3) = "fmFormat": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "SelectMerge": btnArr(iBtn, 2) = "btnSelectMerge_Click": btnArr(iBtn, 3) = "fmFormat": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "数值粘贴": btnArr(iBtn, 2) = "btnPasteValue_Click": btnArr(iBtn, 3) = "fmFormat": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "UnProtectSht": btnArr(iBtn, 2) = "btnUnProtectSht_Click": btnArr(iBtn, 3) = "fmFormat": iBtn = iBtn + 1
        
    'fmOther
    btnArr(iBtn, 1) = "切换引用": btnArr(iBtn, 2) = "btnQieHuanYinYong_Click": btnArr(iBtn, 3) = "fmOther": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "断开外部链接": btnArr(iBtn, 2) = "btnBreakLink_Click": btnArr(iBtn, 3) = "fmOther": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "Unload Me": btnArr(iBtn, 2) = "btnUnLoad_Click": btnArr(iBtn, 3) = "fmOther": iBtn = iBtn + 1
    btnArr(iBtn, 1) = "关闭": btnArr(iBtn, 2) = "btnClose_Click": btnArr(iBtn, 3) = "fmOther": iBtn = iBtn + 1
    
    
    iBtn = iBtn - 1
    ReDim mc(1 To iBtn) As New MyControl
                    
    For i = 1 To iBtn
        Select Case btnArr(i, 3)
            Case "fmFormat"
                Set btn = fmFormat.Controls.Add("Forms.CommandButton.1")
            Case "fmJiaGe"
                Set btn = fmJiaGe.Controls.Add("Forms.CommandButton.1")
            Case "fmOther"
                Set btn = fmOther.Controls.Add("Forms.CommandButton.1")
        End Select
        Set mc(i).btn = btn
        
        With btn
            If btnArr(i, 1) = "Unload Me" Then .Cancel = True
            .Caption = btnArr(i, 1) & Space(20) & "|"
            .Tag = btnArr(i, 2)
        End With
    Next i
    
    iFrTop = 5
    For Each fr In Me.Controls
        If TypeName(fr) = "Frame" Then
            With fr
                .ForeColor = &H8000000D
                .Width = iWidth * 2 + 8
                .Height = ((.Controls.Count + 1) \ 2) * iHeight + 18
                .Top = iFrTop: iFrTop = iFrTop + .Height + 5
                .Left = iLeft
        '        .BorderStyle = fmBorderStyleSingle
        '        .BorderColor = &H8000000D
                
                iBtnTop = 5
                For iBtn = 0 To .Controls.Count - 1
                    Set btn = .Controls(iBtn)
                    btn.Top = iBtnTop: iBtnTop = iBtnTop + iHeight * (iBtn Mod 2)
                    btn.Width = iWidth
                    btn.Height = iHeight
                    btn.Left = iLeft + iWidth * (iBtn Mod 2)
                Next iBtn
            End With
        End If
    Next fr
    
    
    With Me
        .Height = iFrTop + 25
        .Width = iWidth * 2 + 15
        .Caption = "常用命令"
    End With
    
    
    Set btn = Nothing
    Erase btnArr
End Sub
