Attribute VB_Name = "MFunc"
Option Explicit


Function 退休() As ResultStruct
    '基本医疗保险继续按南昌市城镇困难企业医保政策执行
    Dim d1 As Double
    d1 = 0.08 * MMain.M_医疗计算基数 * 15 * 12
    
    退休.医疗保险 = d1
End Function

Function 内部退养(dTuiXiu As Date) As ResultStruct
    '继续为其缴纳养老保险费企业部分
    '按2018年江西省职工社会平均工资60%为基数缴纳(缴纳28%)，以后年度在此基础上每年增加10%
    Dim d1 As Double, k As Long
    
    Dim dTmp As Date
    dTmp = VBA.DateAdd("d", 1, MMain.D_基准日)
    Do Until dTmp > dTuiXiu
        d1 = d1 + MMain.M_社会平均工资 * (1.1 ^ (VBA.Year(dTmp) - VBA.Year(MMain.D_基准日) - 1))
        dTmp = VBA.DateAdd("m", 1, dTmp)
        k = k + 1
    Loop
    d1 = d1 * 0.28 * 0.6
    
    '基本医疗保险继续按南昌市城镇困难企业医保政策执行
    Dim d3 As Double
    d3 = 0.08 * MMain.M_医疗计算基数 * 25 * 12
    
    '基本生活费按1260元/月（2018年度南昌市最低工资的75%）发放，以后年度不调整，退养时间不足一个月的，按一个月发放
    Dim d2 As Double
    d2 = k * MMain.M_最低工资标准 * 0.75
    
    
    内部退养.养老保险 = d1
    内部退养.医疗保险 = d3
    内部退养.Other = d2
End Function

'只计算最多5年
Function 协议社保(dTuiXiu As Date) As ResultStruct
    '基本养老保险按2018年江西省职工社会平均工资60%为基数缴纳，以后年度在此基础上每年增加10%
    Dim d1 As Double, k As Long
    
    Dim dTmp As Date
    dTmp = VBA.DateAdd("d", 1, MMain.D_基准日)
    Do Until dTmp > dTuiXiu
        d1 = d1 + MMain.M_社会平均工资 * (1.1 ^ (VBA.Year(dTmp) - VBA.Year(MMain.D_基准日) - 1))
        dTmp = VBA.DateAdd("m", 1, dTmp)
        k = k + 1
    Loop
    d1 = d1 * 0.28 * 0.6
    
    
    '基本生活费按1260元/月（2018年度南昌市最低工资的75%）发放，但一次性只发放累计不超过5年的基本生活费
    Dim d2 As Double
    If k > 5 * 12 Then k = 5 * 12
    d2 = k * MMain.M_最低工资标准 * 0.75
    
    '基本医疗保险继续按南昌市城镇困难企业医保政策执行
    Dim d3 As Double
    d3 = 0.08 * MMain.M_医疗计算基数 * 25 * 12
    
    协议社保.养老保险 = d1
    协议社保.医疗保险 = d3
    协议社保.Other = d2
End Function


Function 经济补偿(dWork As Date, dTuiXiu As Date) As ResultStruct
    '按其本人连续工龄每满一年支付1680元（南昌市2018年最低工资标准）经济补偿（最低保障标准），工作不满一年按一年计算，一次性支付给职工本人
    Dim d1 As Double
    Dim dTmp As Date
    
    dTmp = MMain.D_基准日
    If dTmp > dTuiXiu Then dTmp = dTuiXiu
    
    d1 = VBA.CDbl(VBA.DateDiff("yyyy", dWork, dTmp)) * MMain.M_最低工资标准
    If d1 < 0# Then d1 = 0#
    
    Dim d2 As Double
    '一次性补发24个月失业金，标准为1260元/月
    d2 = 24# * 1260#
    
    '基本医疗保险继续按南昌市城镇困难企业医保政策执行
    Dim d3 As Double
    d3 = 0.08 * MMain.M_医疗计算基数 * 25 * 12
    
    经济补偿.医疗保险 = d3
    经济补偿.Other = d1 + d2
End Function

Function 工伤残(dWork As Date, dTuiXiu As Date) As Double
    工伤残 = 0
End Function
Function 未参保(dWork As Date, dTuiXiu As Date) As ResultStruct
    '根据职工在企业的工作年限，按其本人在本企业工龄每满一年支付一个月，但未在企业工作的年限的经济补偿金应与扣除，
    '2007年12月31日前经济补偿金标准为380元/月，2008年1月1日以后的经济补偿金标准为1680元/月，最多不超过十二个月
    Dim dTmp As Date, k As Long
    Dim lResult As Long
    
    dTmp = dWork
    Do Until dTmp >= dTuiXiu
        If dTmp <= #12/31/2007# Then
            lResult = lResult + 380
        Else
            lResult = lResult + MMain.M_最低工资标准
        End If
        
        k = k + 1
        If k >= 12 Then Exit Do
        
        dTmp = VBA.DateAdd("yyyy", 1, dTmp)
    Loop
    
    未参保.Other = VBA.CDbl(lResult)
End Function
Function 抚恤() As ResultStruct
    '每人每月补助240元，死亡职工的配偶和父母一次性测算十年
    抚恤.Other = (1# + 2#) * 10# * 12# * 400#
    
    '未成年子女一次性测算至抚养对象满18周岁止，一次性支付，已享受了相关待遇的，不再重复享受
    
End Function

Function GetBirthrDayFromSFZ(strSFZ As String) As Date
    If VBA.Len(strSFZ) = 15 Then
        GetBirthrDayFromSFZ = VBA.DateSerial(VBA.CInt("19" & VBA.Mid$(strSFZ, 7, 2)), VBA.CInt(VBA.Mid$(strSFZ, 9, 2)), VBA.CInt(VBA.Mid$(strSFZ, 11, 2)))
    ElseIf VBA.Len(strSFZ) = 18 Then
        GetBirthrDayFromSFZ = VBA.DateSerial(VBA.CInt(VBA.Mid$(strSFZ, 7, 4)), VBA.CInt(VBA.Mid$(strSFZ, 11, 2)), VBA.CInt(VBA.Mid$(strSFZ, 13, 2)))
    Else
        GetBirthrDayFromSFZ = #12/31/9999#
    End If
End Function

Function GetXingBieFromSFZ(strSFZ As String) As String
    Dim i As Long
    
    If VBA.Len(strSFZ) = 15 Then
        i = VBA.CInt(VBA.Mid$(strSFZ, 15, 1))
    ElseIf VBA.Len(strSFZ) = 18 Then
        i = VBA.CInt(VBA.Mid$(strSFZ, 17, 1))
    Else
        GetXingBieFromSFZ = ""
        Exit Function
    End If
    
    '男的为奇数，女的为偶数
    If i Mod 2 Then
        GetXingBieFromSFZ = "男"
    Else
        GetXingBieFromSFZ = "女"
    End If
End Function
