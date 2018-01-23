Attribute VB_Name = "M股票"
Option Explicit

'缩写找股票代码
Const URL_GET_CODE As String = "http://suggest3.sinajs.cn/suggest/type=&key="
'返回   var suggestvalue="klww,11,300418,sz300418,昆仑万维,klww,昆仑万维,0";

'股票数据
Const URL_GET_CODE_DATA As String = "http://hq.sinajs.cn/list="
'var hq_str_sz300418="昆仑万维,21.570,21.630,21.590,21.980,21.520,21.590,21.600,8343217,180853667.490,39500,21.590,15700,21.580,8200,21.570,10800,21.560,39100,21.550,31400,21.600,24200,21.610,19400,21.620,60900,21.630,9500,21.640,2018-01-05,15:05:03,00";

'http://blog.csdn.net/simon803/article/details/7784682
'0：”大秦铁路”，股票名字；
'1：”27.55″，今日开盘价；
'2：”27.25″，昨日收盘价；
'3：”26.91″，当前价格；
'4：”27.55″，今日最高价；
'5：”26.20″，今日最低价；
'6：”26.91″，竞买价，即“买一”报价；
'7：”26.92″，竞卖价，即“卖一”报价；
'8：”22114263″，成交的股票数，由于股票交易以一百股为基本单位，所以在使用时，通常把该值除以一百；
'9：”589824680″，成交金额，单位为“元”，为了一目了然，通常以“万元”为成交金额的单位，所以通常把该值除以一万；
'10：”4695″，“买一”申请4695股，即47手；
'11：”26.91″，“买一”报价；
'12：”57590″，“买二”
'13：”26.90″，“买二”
'14：”14700″，“买三”
'15：”26.89″，“买三”
'16：”14300″，“买四”
'17：”26.88″，“买四”
'18：”15100″，“买五”
'19：”26.87″，“买五”
'20：”3100″，“卖一”申报3100股，即31手；
'21：”26.92″，“卖一”报价
'(22, 23), (24, 25), (26,27), (28, 29)分别为“卖二”至“卖四的情况”
'30：”2008-01-11″，日期；
'31：”15:05:32″，时间；

Const URL_GET_HISTORY_DATA As String = "http://market.finance.sina.com.cn/downxls.php?"

Sub vba_main()
    Dim i As Long
    Dim arr()
    Dim i_row As Long
    Dim tmp() As String
    
    ActiveSheet.AutoFilterMode = False
    i_row = Cells(Cells.Rows.Count, 1).End(xlUp).Row
    If i_row < 2 Then MsgBox "没有数据": Exit Sub
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

'根据股票代码，获取指定日期的股票数据
'strDate    日期（"2018-01-01"）
'strSymbol  股票代码
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

'根据股票代码，获取最新的股票数据
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

'根据股票的中文名称、首字母缩写来获取股票代码
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
'                Cells(r, 2).Value = sp(1)       '名称
'                Cells(r, 3).Value = sp(3)       '现价
'
'                Cells(r, 5).Value = sp(4)       '昨收
'                Cells(r, 6).Value = sp(5)       '金开
'                Cells(r, 7).Value = sp(31)      '涨跌
'                Cells(r, 8).Value = sp(32)      '涨跌幅
'                Cells(r, 9).Value = sp(33)      '最高
'                Cells(r, 10).Value = sp(34)     '最低
'                Cells(r, 11).Value = sp(43)     '振幅
'                Cells(r, 12).Value = sp(38)     '换手率
'                Cells(r, 13).Value = sp(39)     '市盈率
'               Cells(r, 14).Value = sp(36)     '成交量
'                Cells(r, 15).Value = sp(37)     '成交量
'               Cells(r, 16).Value = sp(44)     '流通市值
'                Cells(r, 17).Value = sp(45)     '总市值
'                Cells(r, 18).Value = sp(47) '涨停
'                Cells(r, 19).Value = sp(48) '跌停
'
'                 Cells(r, 4).Value = Format(sp(s2), "0000-00-00 00:00:00") '时间
