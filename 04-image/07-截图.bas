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

