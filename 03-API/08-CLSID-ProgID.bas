Attribute VB_Name = "模块1"
Option Explicit


Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const SYNCHRONIZE As Long = &H100000
Private Const KEY_READ As Long = (( _
                                  STANDARD_RIGHTS_READ _
                                  Or KEY_QUERY_VALUE _
                                  Or KEY_ENUMERATE_SUB_KEYS _
                                  Or KEY_NOTIFY) _
                                  And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_NO_MORE_ITEMS As Long = 259&
Private Declare Function RegOpenKeyEx _
                          Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
                              ByVal hKey As Long, _
                              ByVal lpSubKey As String, _
                              ByVal ulOptions As Long, _
                              ByVal samDesired As Long, _
                              ByRef phkResult As Long) As Long
Private Declare Function RegEnumKey _
                          Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
                              ByVal hKey As Long, _
                              ByVal dwIndex As Long, _
                              ByVal lpName As String, _
                              ByVal cbName As Long) As Long
Private Declare Function RegQueryValue _
                          Lib "advapi32.dll" Alias "RegQueryValueA" ( _
                              ByVal hKey As Long, _
                              ByVal lpSubKey As String, _
                              ByVal lpValue As String, _
                              ByRef lpcbValue As Long) As Long
Private Declare Function RegCloseKey _
                          Lib "advapi32.dll" ( _
                              ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
                                    (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
                                         (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As _
                                                                                                                     Long, lpData As Any, lpcbData As Long) As Long
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare Function ProgIDFromCLSID Lib "ole32.dll" (ByRef CLSID As Any, ByRef lplpszProgID As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef pclsid As Any) As Long
Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long

Dim R As Long
Sub TypeLibList()
    Dim R1 As Long
    Dim R2 As Long
    Dim hHK1 As Long
    Dim hHK2 As Long
    Dim hHK3 As Long
    Dim hHK4 As Long
    Dim i As Long
    Dim i2 As Long
    Dim lpPath As String
    Dim lpGUID As String
    Dim lpName As String
    Dim lpValue As String
    Application.ScreenUpdating = False
    Cells.Clear: R = 1: Cells(1, 1).Resize(1, 5) = Split("类型库文件路径\类型库引用名称|CLSID|ProgID|默认名称|CLSID对应的库文件", "|")
    lpPath = String$(128, vbNullChar)
    lpValue = String$(128, vbNullChar)
    lpName = String$(128, vbNullChar)
    lpGUID = String$(128, vbNullChar)
    R1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, "TypeLib", ByVal 0&, KEY_READ, hHK1)
    If R1 = ERROR_SUCCESS Then
        i = 0:
        Do While Not R1 = ERROR_NO_MORE_ITEMS
            R1 = RegEnumKey(hHK1, i, lpGUID, Len(lpGUID))
            If R1 = ERROR_SUCCESS Then
                R2 = RegOpenKeyEx(hHK1, lpGUID, ByVal 0&, KEY_READ, hHK2)
                If R2 = ERROR_SUCCESS Then
                    i2 = 0
                    Do While Not R2 = ERROR_NO_MORE_ITEMS
                        R2 = RegEnumKey(hHK2, i2, lpName, Len(lpName))    '1.0
                        If R2 = ERROR_SUCCESS Then
                            RegQueryValue hHK2, lpName, lpValue, Len(lpValue)
                            RegOpenKeyEx hHK2, lpName, ByVal 0&, KEY_READ, hHK3
                            RegOpenKeyEx hHK3, "0", ByVal 0&, KEY_READ, hHK4
                            RegQueryValue hHK4, "win32", lpPath, Len(lpPath)
                            i2 = i2 + 1
                            Cells(R + 1, 1) = IIf(InStr(lpPath, vbNullChar), Left(lpPath, InStr(lpPath, vbNullChar) - 1), lpPath) & Chr(10) _
                                              & IIf(InStr(lpValue, vbNullChar), Left(lpValue, InStr(lpValue, vbNullChar) - 1), lpValue) & Chr(10)
                            ProgIDFromFile lpPath
                        End If
                    Loop
                End If
            End If
            i = i + 1
        Loop
        RegCloseKey hHK1
        RegCloseKey hHK2
        RegCloseKey hHK3
        RegCloseKey hHK4
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub ProgIDFromFile(TypeLibFile$)
    Dim CLSID As GUID, strProgID$, lpszProgID&
    Dim TLIApp As Object
    Dim TLBInfo As Object
    Dim TypeInf As Object
    Set TLIApp = New TLI.TLIApplication
    Dim ProgID As String
    Dim strCLSID As String
    On Error GoTo Exitpoint
    Set TLBInfo = TLIApp.TypeLibInfoFromFile(TypeLibFile)
    For Each TypeInf In TLBInfo.CoClasses
        ProgID = TypeInf.Name
        strCLSID = TypeInf.GUID
        If CLSIDFromString(StrPtr(strCLSID), CLSID) = 0 Then
            R = R + 1: Cells(R, 2) = strCLSID
            Cells(R, 4) = CLSIDDefaultValue(strCLSID)(0)
            Cells(R, 5) = CLSIDDefaultValue(strCLSID)(1)
            If ProgIDFromCLSID(CLSID, lpszProgID) = 0 Then
                SysReAllocString VarPtr(strProgID), lpszProgID
                Cells(R, 3) = strProgID
            End If
        End If
    Next
Exitpoint:

End Sub

Private Function CLSIDDefaultValue(strCLSID$)
    Dim ret As Long
    Dim key As Long
    Dim length As Long
    Dim temp$(0 To 1)
    ret = RegOpenKey(HKEY_CLASSES_ROOT, "CLSID", key)
    ret = RegOpenKey(key, strCLSID, key)
    '先取数据区的长度
    ret = RegQueryValueEx(key, "", 0, 1, ByVal 0, length)
    '准备数据区
    If length > 0 Then
        Dim buff() As Byte
        ReDim buff(length - 1)
        '读取数据
        ret = RegQueryValueEx(key, "", 0, 1, buff(0), length)
'        Dim val As String
        '去掉末尾的空字符,VB不需要这个
        ReDim Preserve buff(length - 2)
        '转化为VB中的字符串
        temp(0) = StrConv(buff, vbUnicode)
    End If
    ret = RegOpenKey(key, "InprocServer32", key)
    ret = RegQueryValueEx(key, "", 0, 1, ByVal 0, length)
    If length > 0 Then
        
        ReDim buff(length - 1)
        '读取数据
        ret = RegQueryValueEx(key, "", 0, 1, buff(0), length)
        
        '去掉末尾的空字符,VB不需要这个
        ReDim Preserve buff(length - 2)
        '转化为VB中的字符串
        temp(1) = StrConv(buff, vbUnicode)
    End If
    CLSIDDefaultValue = temp
    RegCloseKey (key)
End Function


