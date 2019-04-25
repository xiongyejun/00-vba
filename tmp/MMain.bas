Attribute VB_Name = "MMain"
Option Explicit

Public Const D_基准日 As Date = #12/31/2019#
Public Const M_最低工资标准 As Double = 1680#
Public Const M_社会平均工资 As Double = 7000#



Public Enum ReturnCode
    ErrRT = -1
    SuccessRT = 1
End Enum

Public Enum Pos
    Zero
    编号
    姓名
    部门
    性别
    出生日期
    年龄
    政治面貌
    身份证号
    是否有档案
    职工身份
    参加工作时间
    到本企业时间
    连续工龄
    人员类别
    是否在岗
    农历生日
    生日类别
    离岗时间
    是否退休
    退休时间
    是否解除劳动合同
    解除劳动合同原因
    解除劳动合同时间
    备注
    
    RowStart = 2
    KeyCol = 姓名
    Cols = 备注
End Enum

Type DataStruct
    Src() As Variant
    Rows As Long
    Cols As Long
    
    c As CCount
    Result() As Variant
End Type

Sub vba_main()
    Dim d As DataStruct
    
    If ReturnCode.ErrRT = ReadData(d) Then Exit Sub
    
    On Error GoTo err_handle
    
    
    Exit Sub
err_handle:
    MsgBox Err.Description
End Sub


