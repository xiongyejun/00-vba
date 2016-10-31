Attribute VB_Name = "ģ��2"
Option Explicit


Type RECT_Type
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'��*********************************��ͼ��API����***************************************��
Declare Function GetDesktopWindow Lib "user32" () As Long  '��ô���������Ļ��һ�����ڣ����洰�ڣ����
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long '��ȡָ�����ڵ��豸����
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long '����һ����ض��豸����һ�µ��ڴ��豸����
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long '����һ�����豸�й�λͼ�������ƶ������豸��������
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long 'ѡ���Ķ�������豸�����Ļ�ͼ������ʹ��
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long '��һ��λͼ��һ���豸�������Ƶ���һ��
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long '�ͷ��豸�����Ļ�����DC)������Ӧ�ó���ʹ��
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long 'ɾ��ָ���豸�����Ļ�����DC)

Global Const SRCCOPY = &HCC0020
Global Const CF_BITMAP = 2

Private Declare Sub CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, pclsid As Any)
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PicBmp, riid As Any, ByVal fOwn As Long, lplpvObj As IPicture) As Long

Private Type PicBmp
    size As Long
    type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Function ScreenDump(image_rect As RECT_Type)
    Dim DeskHwnd As Long
    Dim hdc As Long
    Dim hdcMem As Long
    Dim Junk As Long
    Dim Fwidth As Long, Fheight As Long
    Dim hBitmap As Long
    Dim pic As StdPicture
    Dim save_file As String
    
    DeskHwnd = GetDesktopWindow()
    
    Fwidth = image_rect.Right - image_rect.Left
    Fheight = image_rect.Bottom - image_rect.Top
    
    hdc = GetDC(DeskHwnd)
    hdcMem = CreateCompatibleDC(hdc)
    hBitmap = CreateCompatibleBitmap(hdc, Fwidth, Fheight)
    Junk = SelectObject(hdcMem, hBitmap)
    Junk = BitBlt(hdcMem, 0, 0, Fwidth, Fheight, hdc, image_rect.Left, image_rect.Top, SRCCOPY)
    Junk = DeleteDC(hdcMem)
    Junk = ReleaseDC(DeskHwnd, hdc)
    
    Set pic = get_pic_from_bitmap(hBitmap)
    
    save_file = Application.GetSaveAsFilename(InitialFileName:=Format(Now(), "yyyy-mm-ddHHMMSS") & ".jpg", filefilter:="JPEG�ļ�(*.jpg),*.jpg", Title:="���ΪͼƬ")

    SavePicture pic, save_file
End Function

Function get_pic_from_bitmap(bitmap_hwnd As Long) As StdPicture
    Dim IID_IDispatch(15) As Byte
    Dim pic As PicBmp
    
    CLSIDFromString StrPtr("{00020400-0000-0000-C000-000000000046}"), IID_IDispatch(0)
    With pic
        .size = Len(pic)
        .type = 1
        .hBmp = bitmap_hwnd
        .hPal = 0
    End With
    
    OleCreatePictureIndirect pic, IID_IDispatch(0), 1, get_pic_from_bitmap
End Function


GDI

Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCllback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Integer
    type As Long
    Value As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal str As Long, id As GUID) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long

'quality jpg�ļ���������1-100֮�����ֵ����ֵԽ��ͼƬ����Խ��
Function save_image_by_gdi(hbmp As Long, Optional ByVal quality As Long = 80)
    Dim lRes As Long
    Dim lGDIP As Long
    Dim tSI As GdiplusStartupInput
    Dim lBitmap As Long
    Dim file_name As String
    
    tSI.GdiplusVersion = 1
    lRes = GdiplusStartup(lGDIP, tSI, 0)
    
    If lRes = 0 Then
        lRes = GdipCreateBitmapFromHBITMAP(hbmp, 0, lBitmap)
        
        If lRes = 0 Then
            Dim tJpgEncoder As GUID
            Dim tParams As EncoderParameters
            
            '��ʼ����������GUID��ʶ
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
            
            '���ý���������
            tParams.Count = 1
            With tParams.Parameter  'Quality
                '�õ�Quality������GUID��ʶ
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                .NumberOfValues = 1
                .type = 4
                .Value = VarPtr(quality)
            End With
            
            '����ͼ��
            file_name = Application.GetSaveAsFilename(InitialFileName:=Format(Now(), "yyyy-mm-ddHHMMSS") & ".jpg", filefilter:="JPEG�ļ�(*.jpg),*.jpg", Title:="���ΪͼƬ")
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(file_name), tJpgEncoder, tParams)
            
            GdipDisposeImage lBitmap
        End If
        
        GdiplusShutdown lGDIP
    End If
End Function


