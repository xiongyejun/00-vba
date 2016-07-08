Attribute VB_Name = "Ä£¿é1"
Option Explicit

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As Any, ByRef lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IUnknown) As Long

Public Function BitmapToPicture(ByVal hBmp As Long, ByVal fPictureOwnsHandle As Long) As StdPicture

    If (hBmp = 0) Then Exit Function

    Dim oNewPic As IUnknown, tPicConv As PictDesc, IGuid As Guid

    With tPicConv
        .cbSizeofStruct = Len(tPicConv)
        .picType = 1 'vbPicTypeBitmap
        .hImage = hBmp
    End With

    With IGuid
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    OleCreatePictureIndirect tPicConv, IGuid, fPictureOwnsHandle, oNewPic

    Set BitmapToPicture = oNewPic

End Function

Public Function ByteArrayToPicture(ByVal lp As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nLeftPadding As Long, Optional ByVal nTopPadding As Long, Optional ByVal nRightPadding As Long, Optional ByVal nBottomPadding As Long) As StdPicture
    Dim tBMI As BITMAPINFO
    Dim h As Long, hdc As Long, hBmp As Long
    Dim hbr As Long
    Dim r As RECT
    With tBMI.bmiHeader
        .biSize = 40&
        .biWidth = nWidth
        .biHeight = -nHeight
        .biPlanes = 1
        .biBitCount = 8
        .biSizeImage = nWidth * nHeight
        .biClrUsed = 256
    End With
    tBMI.bmiColors(0) = &HFFFFFF
    tBMI.bmiColors(2) = &H808080
    h = GetDC(0)
    hdc = CreateCompatibleDC(h)
    r.Right = nWidth + nLeftPadding + nRightPadding
    r.Bottom = nHeight + nTopPadding + nBottomPadding
    hBmp = CreateCompatibleBitmap(h, r.Right, r.Bottom)
    hBmp = SelectObject(hdc, hBmp)
    hbr = CreateSolidBrush(vbWhite)
    FillRect hdc, r, hbr
    DeleteObject hbr
    StretchDIBits hdc, nLeftPadding, nTopPadding, nWidth, nHeight, 0, 0, nWidth, nHeight, ByVal lp, tBMI, 0, 13369376 'vbSrcCopy
    hBmp = SelectObject(hdc, hBmp)
    DeleteDC hdc
    ReleaseDC 0, h
    Set ByteArrayToPicture = BitmapToPicture(hBmp, 1)
End Function

