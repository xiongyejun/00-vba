Attribute VB_Name = "cmd"
Option Explicit

Sub test_cmd()
    Dim ws As Object
    Dim ws_exec As Object
    Dim str As String
    Dim tmp
    
    Set ws = CreateObject("Wscript.Shell")

    
'    ws.Run "cscript", 0
    Set ws_exec = ws.Exec("cmd.exe /c dir ""C:\\Documents and Settings\\xyj\\×ÀÃæ"" /b")
'    Set ws_exec = ws.Exec("ipconfig")

    str = ws_exec.StdOut.ReadAll
    tmp = Split(str, Chr(10))
    Cells.Delete
    Range("a1").Resize(UBound(tmp) + 1, 1).Value = Application.WorksheetFunction.Transpose(tmp)
    
    Erase tmp
    Set ws_exec = Nothing
    Set ws = Nothing

End Sub
