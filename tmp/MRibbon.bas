Attribute VB_Name = "MRibbon"
Option Explicit

Sub RibbonUI_onLoad(Ribbon As IRibbonUI)
    On Error Resume Next
    Ribbon.ActivateTab "TabID"
End Sub

Sub rbCountMain(control As IRibbonControl)
    Call CountMain
End Sub
