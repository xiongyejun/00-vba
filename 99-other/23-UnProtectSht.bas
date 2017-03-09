Attribute VB_Name = "模块1"
Option Explicit

Sub UnProtectSht()
    Dim sht As Worksheet
 
    For Each sht In Sheets
        With sht
            .Protect DrawingObjects:=True, Contents:=True, AllowFiltering:=True
            .Protect DrawingObjects:=False, Contents:=True, AllowFiltering:=True
            .Protect DrawingObjects:=True, Contents:=True, AllowFiltering:=True
            .Unprotect
       End With
    Next
    
    Unload btn.Parent.Parent
End Sub

Sub GetShtProtectPWD()
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    Dim i7 As Integer, i8 As Integer, i9 As Integer
    Dim i10 As Integer, i11 As Integer, i12 As Integer
    Dim str As String
    
    On Error Resume Next
    If ActiveSheet.ProtectContents = False Then
        MsgBox "该工作表没有保护密码！"
        Exit Sub
    End If
    
    For i1 = 65 To 66: For i2 = 65 To 66: For i3 = 65 To 66
        For i4 = 65 To 66: For i5 = 65 To 66: For i6 = 65 To 66
            For i7 = 65 To 66: For i8 = 65 To 66: For i9 = 65 To 66
                For i10 = 65 To 66: For i11 = 65 To 66: For i12 = 32 To 126
                    
                    str = Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) _
                    & Chr(i6) & Chr(i7) & Chr(i8) & Chr(i9) & Chr(i10) & Chr(i11) & Chr(i12)
                    ActiveSheet.Unprotect str
                    
                    If ActiveSheet.ProtectContents = False Then
                        MsgBox "已经解除了工作表保护！"
                        Debug.Print str
                        Exit Sub
                    End If
                    
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
End Sub
                

