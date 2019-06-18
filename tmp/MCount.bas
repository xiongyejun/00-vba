Attribute VB_Name = "MCount"
Option Explicit

'序号    人员类别    条件
'1   退休人员
'2   在册人员
'2.1 内部退养    退休年龄不足5年（含5年）的职工或连续工龄男满30、女满25年
'2.2 协议社保    退休年龄5年以上10年以内（含10年）
'2.3 解除劳动合同终止劳动关系领取经济补偿金  退休年龄10年以上（不含10年）
'2.4 工伤残在职职工安置
'2.5 未参保人员
'3   死亡人员
'4   调出人员
'5   除名人员
'注:男60岁退休、女50岁退休


'按具体类别计算
'1、大集体


'2、小集体
'   全部按照【2.3 解除劳动合同终止劳动关系领取经济补偿金】


Private Enum 补偿类别
    退休
    在册
    内部退养
    协议社保
    经济补偿
    工伤残
    未参保
    抚恤
    
    Other
End Enum

Private Enum ResultEnum
    退休时间
    补偿类别
    
    养老保险
    医疗保险
    Other
    合计
    
    Cols
End Enum

Private Enum 性别
    Men
    Women
End Enum


Sub CountMain()
    Dim d As DataStruct
    
    If ReadData(d) = ErrRT Then Exit Sub
    If GetResult(d) = ErrRT Then Exit Sub
    
    MsgBox "OK"
End Sub

Private Function ReadData(d As DataStruct) As ReturnCode
    Sheet1.Activate
    ActiveSheet.AutoFilterMode = False
    d.Rows = Cells(Cells.Rows.count, Pos.KeyCol).End(xlUp).Row
    If d.Rows < Pos.RowStart Then
        MsgBox "没有数据"
        ReadData = ReturnCode.ErrRT
        Exit Function
    End If
    d.Src = Cells(1, 1).Resize(d.Rows, Pos.Cols).value
     
    ReadData = ReturnCode.SuccessRT
End Function
Private Function GetResult(d As DataStruct) As ReturnCode
    Dim i As Long
    Dim retTmp As ResultStruct
    ReDim d.Result(1 To d.Rows, 1 To ResultEnum.Cols) As Variant
    
    For i = Pos.RowStart To d.Rows
        GetTuiXiuDate d, i
        Get补偿类别 d, i
        
'    If i = 3 Then Stop

        If VBA.IsDate(d.Src(i, Pos.参加工作时间)) Then
            Select Case VBA.CStr(d.Result(i, ResultEnum.补偿类别 + 1)) ', VbGet, VBA.CDate(d.Src(i, Pos.参加工作时间)), VBA.CDate(d.Result(i, ResultEnum.退休时间 + 1)))
            Case "退休"
                retTmp = MFunc.退休()
            Case "抚恤"
                retTmp = MFunc.抚恤()
                
                
            Case "内部退养"
                retTmp = MFunc.内部退养(VBA.CDate(d.Result(i, ResultEnum.退休时间 + 1)))
            Case "协议社保"
                retTmp = MFunc.协议社保(VBA.CDate(d.Result(i, ResultEnum.退休时间 + 1)))
                   
                   
            Case "经济补偿"
                retTmp = MFunc.经济补偿(VBA.CDate(d.Src(i, Pos.参加工作时间)), VBA.CDate(d.Result(i, ResultEnum.退休时间 + 1)))
            Case "未参保"
                retTmp = MFunc.未参保(VBA.CDate(d.Src(i, Pos.参加工作时间)), VBA.CDate(d.Result(i, ResultEnum.退休时间 + 1)))
                        
            End Select
            
            d.Result(i, ResultEnum.养老保险 + 1) = retTmp.养老保险
            If VBA.CStr(d.Src(i, Pos.职工身份)) <> "小集体2" Then d.Result(i, ResultEnum.医疗保险 + 1) = retTmp.医疗保险
            d.Result(i, ResultEnum.Other + 1) = retTmp.Other
            d.Result(i, ResultEnum.合计 + 1) = "=SUM(RC[-3]:RC[-1])"
        End If
    Next
    
    d.Result(1, ResultEnum.养老保险 + 1) = "养老保险"
    d.Result(1, ResultEnum.医疗保险 + 1) = "医疗保险"
    d.Result(1, ResultEnum.Other + 1) = "其他"
    d.Result(1, ResultEnum.补偿类别 + 1) = "补偿类别"
    d.Result(1, ResultEnum.退休时间 + 1) = "正常退休时间"
    d.Result(1, ResultEnum.合计 + 1) = "合计"
    
    Cells(1, Pos.Cols + 1).Resize(d.Rows, ResultEnum.Cols).value = d.Result
    
    
    GetResult = SuccessRT
End Function

Function Get补偿类别(d As DataStruct, iRow As Long) As ReturnCode
    Dim str As String
    Dim dTmp As Date
    Dim retStr As String
    
    If d.Src(iRow, Pos.职工身份) <> "小集体2" Then
        str = VBA.CStr(d.Src(iRow, Pos.人员类别))
        
        If str = "退休人员" Then
            retStr = "退休"
        ElseIf str = "死亡人员" Then
            retStr = "抚恤"
        ElseIf str = "调出人员" Or str = "除名人员" Then
            retStr = "未参保"
        Else '在册人员
            
            If MMain.D_基准日 >= VBA.CDate(d.Result(iRow, ResultEnum.退休时间 + 1)) Then
                retStr = "退休"
                
            ElseIf d.Src(iRow, Pos.性别) = "男" And VBA.Val(d.Src(iRow, Pos.连续工龄)) >= 30 Then
                retStr = "内部退养"
            ElseIf d.Src(iRow, Pos.性别) = "女" And VBA.Val(d.Src(iRow, Pos.连续工龄)) >= 25 Then
                retStr = "内部退养"
                
            ElseIf VBA.DateAdd("yyyy", 5, MMain.D_基准日) >= VBA.CDate(d.Result(iRow, ResultEnum.退休时间 + 1)) Then
                retStr = "内部退养"
            ElseIf VBA.DateAdd("yyyy", 10, MMain.D_基准日) >= VBA.CDate(d.Result(iRow, ResultEnum.退休时间 + 1)) Then
                retStr = "协议社保"
            Else
               retStr = "经济补偿"
            End If
        
        End If
    Else
        retStr = "经济补偿"
    End If
    d.Result(iRow, ResultEnum.补偿类别 + 1) = retStr
    
    Get补偿类别 = SuccessRT
End Function

'退休日期
Private Function GetTuiXiuDate(d As DataStruct, iRow As Long) As ReturnCode
    Dim xb As 性别
    
    If VBA.IsDate(d.Src(iRow, Pos.出生日期)) Then
        If d.Src(iRow, Pos.性别) = "男" Then
            xb = Men
        Else
            xb = Women
        End If
            
        d.Result(iRow, ResultEnum.退休时间 + 1) = VBA.DateAdd("yyyy", 50 - 10 * (xb = Men), VBA.CDate(d.Src(iRow, Pos.出生日期)))
    Else
    
    End If
    
    GetTuiXiuDate = SuccessRT
End Function

Sub TestTmp()
    Dim d As DataStruct
    
    If ReadData(d) = ErrRT Then Exit Sub
    
    Dim i As Long
    Dim dTmp As Date
    
    For i = 2 To d.Rows
        If Not VBA.IsDate(d.Src(i, Pos.出生日期)) Then
            dTmp = MFunc.GetBirthrDayFromSFZ(VBA.CStr(d.Src(i, Pos.身份证号)))
            If dTmp <> #12/31/9999# Then
                With Cells(i, Pos.出生日期)
                    .value = dTmp
                    .Interior.Color = 255
                End With
            End If
            
        End If
    Next
    
End Sub