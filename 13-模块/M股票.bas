Attribute VB_Name = "M��Ʊ"
Option Explicit

'��д�ҹ�Ʊ����
Const URL_GET_CODE As String = "http://suggest3.sinajs.cn/suggest/type=&key="
'����   var suggestvalue="klww,11,300418,sz300418,������ά,klww,������ά,0";

'��Ʊ����
Const URL_GET_CODE_DATA As String = "http://hq.sinajs.cn/list="
'var hq_str_sz300418="������ά,21.570,21.630,21.590,21.980,21.520,21.590,21.600,8343217,180853667.490,39500,21.590,15700,21.580,8200,21.570,10800,21.560,39100,21.550,31400,21.600,24200,21.610,19400,21.620,60900,21.630,9500,21.640,2018-01-05,15:05:03,00";

'http://blog.csdn.net/simon803/article/details/7784682
'0����������·������Ʊ���֣�
'1����27.55�壬���տ��̼ۣ�
'2����27.25�壬�������̼ۣ�
'3����26.91�壬��ǰ�۸�
'4����27.55�壬������߼ۣ�
'5����26.20�壬������ͼۣ�
'6����26.91�壬����ۣ�������һ�����ۣ�
'7����26.92�壬�����ۣ�������һ�����ۣ�
'8����22114263�壬�ɽ��Ĺ�Ʊ�������ڹ�Ʊ������һ�ٹ�Ϊ������λ��������ʹ��ʱ��ͨ���Ѹ�ֵ����һ�٣�
'9����589824680�壬�ɽ�����λΪ��Ԫ����Ϊ��һĿ��Ȼ��ͨ���ԡ���Ԫ��Ϊ�ɽ����ĵ�λ������ͨ���Ѹ�ֵ����һ��
'10����4695�壬����һ������4695�ɣ���47�֣�
'11����26.91�壬����һ�����ۣ�
'12����57590�壬�������
'13����26.90�壬�������
'14����14700�壬��������
'15����26.89�壬��������
'16����14300�壬�����ġ�
'17����26.88�壬�����ġ�
'18����15100�壬�����塱
'19����26.87�壬�����塱
'20����3100�壬����һ���걨3100�ɣ���31�֣�
'21����26.92�壬����һ������
'(22, 23), (24, 25), (26,27), (28, 29)�ֱ�Ϊ���������������ĵ������
'30����2008-01-11�壬���ڣ�
'31����15:05:32�壬ʱ�䣻

Const URL_GET_HISTORY_DATA As String = "http://market.finance.sina.com.cn/downxls.php?"

Sub vba_main()
    Dim i As Long
    Dim arr()
    Dim i_row As Long
    Dim tmp() As String
    
    ActiveSheet.AutoFilterMode = False
    i_row = Cells(Cells.Rows.Count, 1).End(xlUp).Row
    If i_row < 2 Then MsgBox "û������": Exit Sub
    arr = Range("A1:A" & VBA.CStr(i_row)).Value
    For i = 2 To i_row
        tmp = GetCodeData(GetCode(VBA.CStr(arr(i, 1))))
        If VBA.IsArray(tmp) Then Cells(i, 2).Resize(1, UBound(tmp) - LBound(tmp) + 1) = tmp
    Next
    
    On Error GoTo err_handle
    
    
    Exit Sub
err_handle:
    MsgBox Err.Description
End Sub

Sub testas()
    GetHistoryData "2018-01-01", "sz300418"
End Sub

'���ݹ�Ʊ���룬��ȡָ�����ڵĹ�Ʊ����
'strDate    ���ڣ�"2018-01-01"��
'strSymbol  ��Ʊ����
Private Function GetHistoryData(strDate As String, strSymbol As String) As String
    'date=2018-01-01&symbol=sz300418
    Dim str As String
    
    Dim obj_http As Object
    
    Set obj_http = CreateObject("WinHttp.WinHttpRequest.5.1") 'CreateObject("MSXML2.XMLHTTP")
    With obj_http
        .Open "GET", URL_GET_HISTORY_DATA & "date=" & strDate & "&symbol=" & strSymbol, False
        .setRequestHeader "Content-Type", "Application/x-www-form-urlencoded"
        .send
        str = VBA.StrConv(.ResponseBody, vbUnicode)
    End With
    
    Set obj_http = Nothing
    
    Debug.Print VBA.Left$(str, 133)
    
End Function

'���ݹ�Ʊ���룬��ȡ���µĹ�Ʊ����
Private Function GetCodeData(strCode As String) As String()
    Dim str As String
    
    str = GetHtml(URL_GET_CODE_DATA & strCode)
    Dim tmp
    tmp = VBA.Split(str, ",")
    
    Dim i As Long
    Dim arr() As String
    
    If UBound(tmp) > 1 Then
        ReDim arr(UBound(tmp) - 1) As String
        For i = 1 To UBound(tmp)
            arr(i - 1) = tmp(i)
        Next i
        arr(i - 2) = VBA.Split(arr(i - 2), """;")(0)
        GetCodeData = arr
    End If
    
End Function

'���ݹ�Ʊ���������ơ�����ĸ��д����ȡ��Ʊ����
Private Function GetCode(strKey As String) As String
    Dim str As String
    
    str = GetHtml(URL_GET_CODE & strKey)
    
    Dim tmp
    tmp = VBA.Split(str, ",")
    If UBound(tmp) > 3 Then
        GetCode = tmp(3)
    End If
End Function

Private Function GetHtml(strURL As String) As String
    Dim obj_http As Object
    
    Set obj_http = CreateObject("WinHttp.WinHttpRequest.5.1") 'CreateObject("MSXML2.XMLHTTP")
    With obj_http
        .Open "GET", strURL, False
        .setRequestHeader "Content-Type", "Application/x-www-form-urlencoded"
        .send
        GetHtml = .responsetext
    End With
    
    Set obj_http = Nothing
End Function

'
'        If Val(dm) < 600000 Then
'            Url = "http://qt.gtimg.cn/q=sz" & dm
'        Else
'            Url = "http://qt.gtimg.cn/q=sh" & dm
'        End If

'            sp = Split(.responsetext, "~")
'            If UBound(sp) > 1 Then
'                Cells(r, 2).Value = sp(1)       '����
'                Cells(r, 3).Value = sp(3)       '�ּ�
'
'                Cells(r, 5).Value = sp(4)       '����
'                Cells(r, 6).Value = sp(5)       '��
'                Cells(r, 7).Value = sp(31)      '�ǵ�
'                Cells(r, 8).Value = sp(32)      '�ǵ���
'                Cells(r, 9).Value = sp(33)      '���
'                Cells(r, 10).Value = sp(34)     '���
'                Cells(r, 11).Value = sp(43)     '���
'                Cells(r, 12).Value = sp(38)     '������
'                Cells(r, 13).Value = sp(39)     '��ӯ��
'               Cells(r, 14).Value = sp(36)     '�ɽ���
'                Cells(r, 15).Value = sp(37)     '�ɽ���
'               Cells(r, 16).Value = sp(44)     '��ͨ��ֵ
'                Cells(r, 17).Value = sp(45)     '����ֵ
'                Cells(r, 18).Value = sp(47) '��ͣ
'                Cells(r, 19).Value = sp(48) '��ͣ
'
'                 Cells(r, 4).Value = Format(sp(s2), "0000-00-00 00:00:00") 'ʱ��
