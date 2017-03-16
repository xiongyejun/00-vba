Attribute VB_Name = "BarCode"
'http://club.excelhome.net/thread-1297553-1-1.html

Private Sub CommandButton1_Click()
    Dim s As String
    s = "*" & TextBox1.Text & "*"
    If TextBox1.Text <> "" Then
        Image1.Picture = BitToPic(GenCode2Bitmap2(Code2Bin(s), 2, 80, 5), 1)
    End If
End Sub


'声明API函数
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As GUID, _
        ByVal fPictureOwnsHandle As Long, IPic As IUnknown) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'定义矩形
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'定义PICTDESC结构
Private Type PictDesc
    cbSizeOfStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
'定义GUID
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
'变量hBmp
Private hBmp As Long
'构造39条形码逻辑态-Gary Fang
Public Function Code2Bin(ByVal str39Code As String) As String
    Dim K As Integer
    Dim strAux As String
    Dim strExit As String
    Dim strCode As String
    
    strExit = ""
    str39Code = Trim(str39Code)
    strAux = str39Code
    
    For K = 1 To Len(str39Code)
        Select Case Mid(strAux, K, 1)
        Case 0
            strExit = strExit & "000110100" & "0"
        Case 1
            strExit = strExit & "100100001" & "0"
        Case 2
            strExit = strExit & "001100001" & "0"
        Case 3
            strExit = strExit & "101100000" & "0"
        Case 4
            strExit = strExit & "000110001" & "0"
        Case 5
            strExit = strExit & "100110000" & "0"
        Case 6
            strExit = strExit & "001110000" & "0"
        Case 7
            strExit = strExit & "000100101" & "0"
        Case 8
            strExit = strExit & "100100100" & "0"
        Case 9
            strExit = strExit & "001100100" & "0"
        Case "A"
            strExit = strExit & "100001001" & "0"
        Case "B"
            strExit = strExit & "001001001" & "0"
        Case "C"
            strExit = strExit & "101001000" & "0"
        Case "D"
            strExit = strExit & "000011001" & "0"
        Case "E"
            strExit = strExit & "100011000" & "0"
        Case "F"
            strExit = strExit & "001011000" & "0"
        Case "G"
            strExit = strExit & "000001101" & "0"
        Case "H"
            strExit = strExit & "100001100" & "0"
        Case "I"
            strExit = strExit & "001001100" & "0"
        Case "J"
            strExit = strExit & "000011100" & "0"
        Case "K"
            strExit = strExit & "100000011" & "0"
        Case "L"
            strExit = strExit & "001000011" & "0"
        Case "M"
            strExit = strExit & "101000010" & "0"
        Case "N"
            strExit = strExit & "000010011" & "0"
        Case "O"
            strExit = strExit & "100010010" & "0"
        Case "P"
            strExit = strExit & "001010010" & "0"
        Case "Q"
            strExit = strExit & "000000111" & "0"
        Case "R"
            strExit = strExit & "100000110" & "0"
        Case "S"
            strExit = strExit & "001000110" & "0"
        Case "T"
            strExit = strExit & "000010110" & "0"
        Case "U"
            strExit = strExit & "110000001" & "0"
        Case "V"
            strExit = strExit & "011000001" & "0"
        Case "W"
            strExit = strExit & "111000000" & "0"
        Case "X"
            strExit = strExit & "010010001" & "0"
        Case "Y"
            strExit = strExit & "110010000" & "0"
        Case "Z"
            strExit = strExit & "011010000" & "0"
        Case "-"
            strExit = strExit & "010000101" & "0"
        Case "%"
            strExit = strExit & "000101010" & "0"
        Case "$"
            strExit = strExit & "010101000" & "0"
        Case "*"
            strExit = strExit & "010010100" & "0"
        Case "+"
            strExit = strExit & "010001010" & "0"
        Case "/"
            strExit = strExit & "010100010" & "0"
        Case "."
            strExit = strExit & "110000100" & "0"
        End Select
    Next
    Code2Bin = strExit
End Function

 '生成固定长宽的条形码Bitmap-Gary Fang
Public Function GenCode2Bitmap(strCode As String, nWidth As Long, nHeight As Long, Margin As Long) As Long
    Dim h As Long, hdc As Long
    Dim hBr As Long
    Dim R As RECT
    Dim DestWidth As Long, DestHeight As Long
    
    DestWidth = nWidth
    DestHeight = nHeight
    R.Top = Margin
    R.Bottom = nHeight - Margin
    
    Dim i As Integer
    Dim Count As Integer
    Count = Len(strCode)
    Dim WCount, NCount As Integer
    WCount = 0
    
    For i = 1 To Count
        If Mid(strCode, i, 1) = "1" Then
            WCount = WCount + 1
        End If
    Next
    
    NCount = Count - WCount
    Dim Unit As Integer
    Unit = Int((DestWidth - 2 * Margin) / (2 * WCount + NCount))
    h = GetDC(0)
    hdc = CreateCompatibleDC(h)
    hBmp = CreateCompatibleBitmap(h, nWidth, nHeight)
    hBmp = SelectObject(hdc, hBmp)
    
    Dim Screen As RECT
    Screen.Left = 0
    Screen.Top = 0
    Screen.Right = DestWidth
    Screen.Bottom = DestHeight
    R.Left = Margin
    '设置背景颜色为白色
    hBr = CreateSolidBrush(RGB(255, 255, 255))
    SelectObject hdc, hBr
    'Rectangle hdc, 0, 0, DestWidth, DestHeight
    FillRect hdc, Screen, hBr
    
    For i = 1 To Len(strCode)
        If i Mod 2 <> 0 Then
            hBr = CreateSolidBrush(RGB(0, 0, 0))
        Else
            hBr = CreateSolidBrush(RGB(255, 255, 255))
        End If
        
        If Mid(strCode, i, 1) = "1" Then
            R.Right = R.Left + 2 * Unit
            SelectObject hdc, hBr
            'Rectangle hdc, R.Left, R.Top, R.Right, R.Bottom
            FillRect hdc, R, hBr
            R.Left = R.Right
        Else
            R.Right = R.Left + Unit
            SelectObject hdc, hBr
            'Rectangle hdc, R.Left, R.Top, R.Right, R.Bottom
            FillRect hdc, R, hBr
            R.Left = R.Right
        End If
    Next
    
    hBmp = SelectObject(hdc, hBmp)
    Call DeleteDC(hdc)
    Call ReleaseDC(0, h)
    GenCode2Bitmap = hBmp
End Function

'生成自由长度但固定高度的条形码Bitmap-Gary Fang
Public Function GenCode2Bitmap2(strCode As String, Unit As Integer, nHeight As Long, Margin As Long) As Long
    Dim h As Long, hdc As Long
    DestHeight = nHeight
    Dim i As Integer
    Dim Count As Integer
    Count = Len(strCode)
    Dim WCount, NCount As Integer
    WCount = 0
    For i = 1 To Count
        If Mid(strCode, i, 1) = "1" Then
            WCount = WCount + 1
        End If
    Next
    NCount = Count - WCount
    DestWidth = NCount * Unit + 2 * Unit * WCount + 2 * Margin
    
    h = GetDC(0)
    hdc = CreateCompatibleDC(h)
    hBmp = CreateCompatibleBitmap(h, DestWidth, DestHeight)
    hBmp = SelectObject(hdc, hBmp)
    
    Dim Screen As RECT
    Screen.Left = 0
    Screen.Top = 0
    Screen.Right = DestWidth
    Screen.Bottom = DestHeight
    Dim R As RECT
    R.Top = Margin
    R.Bottom = DestHeight - Margin
    R.Left = Margin
    '设置背景颜色为白色
    hBr = CreateSolidBrush(RGB(255, 255, 255))
    SelectObject hdc, hBr
    FillRect hdc, Screen, hBr
    'Rectangle hdc, 0, 0, DestWidth, DestHeight
    
    For i = 1 To Len(strCode)
        If i Mod 2 <> 0 Then
            hBr = CreateSolidBrush(RGB(0, 0, 0))
        Else
            hBr = CreateSolidBrush(RGB(255, 255, 255))
        End If
        
        If Mid(strCode, i, 1) = "1" Then
            R.Right = R.Left + 2 * Unit
            SelectObject hdc, hBr
            FillRect hdc, R, hBr
            'Rectangle hdc, R.Left, R.Top, R.Right, R.Bottom
            R.Left = R.Right
        Else
            R.Right = R.Left + Unit
            SelectObject hdc, hBr
            'Rectangle hdc, R.Left, R.Top, R.Right, R.Bottom
            FillRect hdc, R, hBr
            R.Left = R.Right
        End If
    Next
    
    hBmp = SelectObject(hdc, hBmp)
    Call DeleteDC(hdc)
    Call ReleaseDC(0, h)
    GenCode2Bitmap2 = hBmp
End Function

'BitToPic函数来源于公开网站，非本人原创-Gary Fang
Public Function BitToPic(ByVal hBmp As Long, ByVal fPictureOwnsHandle As Long) As StdPicture
    If (hBmp = 0) Then Exit Function
    
    Dim oNewPic As IUnknown, tPicConv As PictDesc, IGuid As GUID
    With tPicConv
        .cbSizeOfStruct = Len(tPicConv)
        .picType = 1
        .hImage = hBmp
    End With
    
    With IGuid
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    OleCreatePictureIndirect tPicConv, IGuid, fPictureOwnsHandle, oNewPic
    Set BitToPic = oNewPic
End Function

