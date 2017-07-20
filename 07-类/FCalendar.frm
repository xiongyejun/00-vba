VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FCalendar 
   Caption         =   "日期"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   OleObjectBlob   =   "FCalendar.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "FCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents btnOK As MSForms.CommandButton
Attribute btnOK.VB_VarHelpID = -1
Private WithEvents btnToday As MSForms.CommandButton
Attribute btnToday.VB_VarHelpID = -1

'6行7列的day
Private ArrC_cal(5, 6) As CCalendar

'
Private d_Date As Date
Private i_Year As Integer
Private i_Month As Integer
Private i_Day As Integer
'
'用来记录上一个选中的day的下标，用来恢复底色
Private i_ArrRow As Integer
Private i_ArrCol As Integer

'日期
Property Let TheDate(Value As Date)
    d_Date = Value
    
    Me.lbDate.Caption = VBA.Format(d_Date, "yyyy年mm月dd日")
    InitDay
End Property
Property Get TheDate() As Date
    TheDate = d_Date
End Property
'年份
Property Let Year(Value As Integer)
    i_Year = Value
    
    Me.TheDate = VBA.DateSerial(i_Year, i_Month, i_Day)
End Property
'月
Property Let Month(Value As Integer)
    i_Month = Value
    
    Me.TheDate = VBA.DateSerial(i_Year, i_Month, i_Day)
End Property
'日
Property Let Day(Value As Integer)
    i_Day = Value
    
    Me.TheDate = VBA.DateSerial(i_Year, i_Month, i_Day)
End Property
'恢复上一个按钮的背景色
Property Let ArrRow(Value As Integer)
    i_ArrRow = Value
End Property
Property Let ArrCol(Value As Integer)
    i_ArrCol = Value
End Property
Sub SetBtnBackColor()
    ArrC_cal(i_ArrRow, i_ArrCol).btn.BackColor = &H8000000F
End Sub


Private Sub btnOK_Click()
    Me.Hide
End Sub

Private Sub btnToday_Click()
    cbMonth.ListIndex = VBA.Month(Date) - 1
    cbYear.Value = VBA.Year(Date)
    Me.Day = VBA.Day(Date)
    
    Dim i As Long, j As Long
    For i = 0 To UBound(ArrC_cal, 1)
        For j = 0 To UBound(ArrC_cal, 2)
            If ArrC_cal(i, j).btn.Caption = VBA.CStr(i_Day) Then
                SetBtnBackColor
                ArrC_cal(i, j).btn.BackColor = &H8000000C
                
                i_ArrRow = i
                i_ArrCol = j
                Exit For
            End If
        Next j
    Next i
    
    
End Sub

Private Sub UserForm_Initialize()
    Dim btn As MSForms.CommandButton
    Dim i As Long, j As Long
    Dim i_left As Integer, i_top As Integer
    Dim btn_width As Integer
    Const BTN_HEIGHT As Integer = 20
    Dim lb As MSForms.Label
    
    btn_width = (Me.cbYear.Left + Me.cbYear.Width) / 7
    i_top = Me.lbDate.Height + Me.lbDate.Top + 5
    
    '添加标题
    For i = #7/2/2017# To #7/2/2017# + 6
        Set lb = Me.Controls.Add("Forms.Label.1")
        With lb
            .Left = 5 + (i - #7/2/2017#) * btn_width
            .Top = i_top
            .Width = btn_width
            .Height = 12
            .TextAlign = fmTextAlignCenter
            .Caption = VBA.Format(i, "AAA")
            .BorderStyle = fmBorderStyleSingle
        End With
        
        If i = #7/2/2017# Or i = #7/2/2017# + 6 Then lb.ForeColor = &HFF&
    Next
    i_top = i_top + lb.Height

    For i = 0 To UBound(ArrC_cal, 1)
        For j = 0 To UBound(ArrC_cal, 2)
            Set btn = Me.Controls.Add("Forms.CommandButton.1")
            
            With btn
                .Left = 5 + j * btn_width
                .Top = i_top + i * BTN_HEIGHT
                .Width = btn_width
                .Height = BTN_HEIGHT
                .Tag = VBA.CStr(i) & "、" & VBA.CStr(j)
            End With
            If j = 0 Or j = 6 Then btn.ForeColor = &HFF&
            Set ArrC_cal(i, j) = New CCalendar
            Set ArrC_cal(i, j).btn = btn
        Next j
    Next i
    
    i_top = i_top + i * BTN_HEIGHT
    Set btnOK = btnAdd(Me, "确定", 5, i_top, (Me.cbYear.Left + Me.cbYear.Width - 3) / 2)
    Set btnToday = btnAdd(Me, "今天", 5 + btnOK.Width, i_top, btnOK.Width)
    
    Me.Height = i_top + btnOK.Height + 25

    Me.Day = VBA.Day(Date)
    CBMonthItem
    CBYearItem
    
    Set lb = Nothing
    Set btn = Nothing
End Sub
'初始化day按钮控件
Function InitDay()
    Dim i_start_week As Integer
    Dim end_day As Integer
    Dim pre_end_day As Integer '上月的最后day
    Dim i As Long, j As Long
    Dim k As Long
    
    i_start_week = VBA.Weekday(VBA.DateSerial(i_Year, i_Month, 1), vbSunday)
    end_day = VBA.Day(VBA.DateSerial(i_Year, i_Month + 1, 0))
    pre_end_day = VBA.Day(VBA.DateSerial(i_Year, i_Month, 0))
    k = 0
    For i = 0 To UBound(ArrC_cal, 1)
        For j = 0 To UBound(ArrC_cal, 2)
            k = k + 1
            If k < i_start_week Then
                ArrC_cal(i, j).btn.Caption = VBA.CStr((k - i_start_week) Mod end_day + 1 + pre_end_day)
            Else
                ArrC_cal(i, j).btn.Caption = VBA.CStr((k - i_start_week) Mod end_day + 1)
            End If
            
            ArrC_cal(i, j).btn.Enabled = (k >= i_start_week And (k - i_start_week + 1) <= end_day)
        Next
    Next
End Function
'月份下拉框元素添加
Function CBMonthItem()
    Dim i As Integer
    
    cbMonth.Clear
    For i = 1 To 12
        cbMonth.AddItem VBA.CStr(i) & "月"
    Next
    
    cbMonth.ListIndex = VBA.Month(Date) - 1
    cbMonth.Style = fmStyleDropDownList
End Function
'年份下拉框元素添加
Function CBYearItem()
    Dim i As Integer
    Dim i_now As Integer
    
    i_now = VBA.Year(Date)
    cbYear.Clear
    For i = 1900 To i_now + 100
        cbYear.AddItem VBA.CStr(i)
    Next
    
    cbYear.Value = i_now
    cbYear.Style = fmStyleDropDownList
End Function

Private Sub cbMonth_Change()
    Me.Month = cbMonth.ListIndex + 1
End Sub

Private Sub cbYear_Change()
    Me.Year = VBA.Val(cbYear.Value)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Me.Hide
    Cancel = 1
End Sub

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
