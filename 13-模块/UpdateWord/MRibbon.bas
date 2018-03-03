Attribute VB_Name = "MRibbon"
Option Explicit

Sub RibbonUI_onLoad(Ribbon As IRibbonUI)
    Ribbon.ActivateTab "TabID"
End Sub

Sub rbUpdateExcel(control As IRibbonControl)
    Call UpdateExcel
End Sub

Sub rbUpdateWord(control As IRibbonControl)
    Call UpdateWordByCom
End Sub

Sub rbAddComToWord(control As IRibbonControl)
    Call AddComToWord
End Sub
