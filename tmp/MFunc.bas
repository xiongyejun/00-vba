Attribute VB_Name = "MFunc"
Option Explicit

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
