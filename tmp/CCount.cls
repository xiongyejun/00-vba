Attribute VB_Name = "CCount"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'用类来做，方面按名称调用

Property Get Test(value As Long) As Long
    Test = 100 + value
End Property

Property Get 退休(dWork As Date, dTuiXiu As Date) As Double
    '基本医疗保险继续按南昌市城镇困难企业医保政策执行
    Dim d1 As Double
    d1 = 0.08 * MMain.M_最低工资标准 * 15 * 12
    
    退休 = d1
End Property

Property Get 内部退养(dWork As Date, dTuiXiu As Date) As Double
    '继续为其缴纳养老保险费企业部分
    '按2018年江西省职工社会平均工资60%为基数缴纳(缴纳28%)，以后年度在此基础上每年增加10%
    Dim d1 As Double, k As Long
    Dim dValue As Double
    dValue = MMain.M_社会平均工资 * 0.6
    
    d1 = dValue
    Dim dTmp As Date
    dTmp = VBA.DateAdd("m", 1, MMain.D_基准日)
    k = 1
    Do Until dTmp > dTuiXiu
        d1 = d1 + dValue
        
        dTmp = VBA.DateAdd("m", 1, dTmp)
        k = k + 1
        If VBA.Month(dTmp) = 12 Then
            dValue = dValue * 1.1
        End If
    Loop
    d1 = d1 * 0.28
    
    '基本医疗保险继续按南昌市城镇困难企业医保政策执行
    Dim d3 As Double
    d3 = 0.08 * MMain.M_最低工资标准 * 25 * 12
    
    '基本生活费按1260元/月（2018年度南昌市最低工资的75%）发放，以后年度不调整，退养时间不足一个月的，按一个月发放
    Dim d2 As Double
    d2 = k * MMain.M_最低工资标准 * 0.75
    
    
    内部退养 = d1 + d2 + d3
End Property

'只计算最多5年
Property Get 协议社保(dWork As Date, dTuiXiu As Date) As Double
    '基本养老保险按2018年江西省职工社会平均工资60%为基数缴纳，以后年度在此基础上每年增加10%
    Dim d1 As Double, k As Long
    Dim dValue As Double
    dValue = MMain.M_社会平均工资 * 0.6
    
    d1 = dValue
    Dim dTmp As Date
    dTmp = VBA.DateAdd("m", 1, MMain.D_基准日)
    k = 1
    Do Until dTmp > dTuiXiu
        d1 = d1 + dValue
        
        dTmp = VBA.DateAdd("m", 1, dTmp)
        k = k + 1
        If VBA.Month(dTmp) = 12 Then
            dValue = dValue * 1.1
        End If
    Loop
    If dTmp > dTuiXiu Then
        d1 = d1 + dValue
        k = k + 1
    End If
    d1 = d1 * 0.28
    
    
    '基本生活费按1260元/月（2018年度南昌市最低工资的75%）发放，但一次性只发放累计不超过5年的基本生活费
    Dim d2 As Double
    If k > 5 * 12 Then k = 5 * 12
    d2 = k * MMain.M_最低工资标准 * 0.75
    
    '基本医疗保险继续按南昌市城镇困难企业医保政策执行
    Dim d3 As Double
    d3 = 0.08 * MMain.M_最低工资标准 * 25 * 12
    
    协议社保 = d1 + d2 + d3
End Property


Property Get 经济补偿(dWork As Date, dTuiXiu As Date) As Double
    '按其本人连续工龄每满一年支付1680元（南昌市2018年最低工资标准）经济补偿（最低保障标准），工作不满一年按一年计算，一次性支付给职工本人
    Dim d1 As Double
    d1 = VBA.CDbl(VBA.DateDiff("yyyy", dWork, MMain.D_基准日)) * MMain.M_最低工资标准
    
    Dim d2 As Double
    '一次性补发24个月失业金，标准为1260元/月
    d2 = 24# * 1260#
    
    '基本医疗保险继续按南昌市城镇困难企业医保政策执行
    Dim d3 As Double
    d3 = 0.08 * MMain.M_最低工资标准 * 25 * 12
    
    经济补偿 = d1 + d2 + d3
End Property

Property Get 工伤残(dWork As Date, dTuiXiu As Date) As Double
    工伤残 = 0
End Property
Property Get 未参保(dWork As Date, dTuiXiu As Date) As Double
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
        If k > 12 Then Exit Do
        
        dTmp = VBA.DateAdd("yyyy", 1, dTmp)
    Loop
    
    未参保 = VBA.CDbl(lResult)
End Property
Property Get 抚恤(dWork As Date, dTuiXiu As Date) As Double
    '每人每月补助240元，死亡职工的配偶和父母一次性测算十年
    抚恤 = (1# + 2#) * 10# * 12# * 400#
    
    '未成年子女一次性测算至抚养对象满18周岁止，一次性支付，已享受了相关待遇的，不再重复享受
    
End Property
