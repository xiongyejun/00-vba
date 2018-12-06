Attribute VB_Name = "MUpdateWorkbook"
Option Explicit

Private Enum Pos
    RowStart = 2
    
    SrcWk = 2
    SrcSht
    SrcRowStart
    SrcColKey
    SrcRng
    
    DesWk
    DesSht
    DesRowStart
    DesColKey
    DesRng
    
    TheType
    BeiZhu
End Enum

Private Enum UpdateType
    rng       '�����ĵ�Ԫ��Ե�Ԫ��ֵ
    ColRelation    '���еĶ�Ӧ��ϵ��һ��һ�еĸ�ֵ
    ColRelationAppend '׷��--������Ĳ�ͬ���ڸ�ֵǰ��Ҫ�������
    Formula     'ֱ��д����ʽ��ȥ
    dic         '��colKey��¼��dic��д��,û���ҵ��ľ�Ϊ��
    dicExists    'ͬ�ϣ������û���ҵ��ģ��ͱ���ԭ��������
    AddName     '����Զ�������---Ŀ�깤������������Ŀ�굥Ԫ���¼�Ķ����ļ����ƣ�����Ҫ���
                'RowStart��ColKey��long���ͣ�����¼
    
    AddNo '������
End Enum
'excel�ļ�֮��ĸ���
Private Type DataStructItem
    wkName As String
    shtName As String
    Action As String        '��¼������str
                            'Rng           ��¼���ǵ�Ԫ���ַ
                            'ColRelation   ��¼�����еĶ�Ӧ��ϵ-ColRelation
                            'Formula       ��¼���ǹ�ʽ��Ŀ�굥Ԫ���¼��Ԫ��Դ��Ԫ��д��ʽ
                            'dic           ��¼����item����
                            
    RowStart As Long        '�����п�ʼ��λ��
    ColKey As Long          '��λ�õ���
End Type

'��¼��������Դ����������Ҫ���棬Ŀ�깤������Ҫ����
Private Type WkType
    wk As Workbook
    bSave As Boolean
    wkName As String
End Type

Private Type DataStruct
    wk As Workbook      '��¼��Ӧ��ϵ��wk
    sht As Worksheet    '��¼��Ӧ��ϵ��sht
    Path As String
    Rows As Long
    Arr() As Variant
    
    Count As Long
    Src() As DataStructItem
    Des() As DataStructItem
    uType() As UpdateType

    dicWk As Object '�ֵ��¼�����������--��ӦArrWk���±�
    ArrWk() As WkType
End Type

'RangeByCol��¼�ж�Ӧ��ϵ�ķָ���
Private Const SPLIT_WORD As String = "��"

Sub UpdateWorkbook()
    Dim d As DataStruct
    
    On Error GoTo err_handle
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set d.wk = ActiveWorkbook
    Set d.sht = ActiveSheet
    d.Path = d.wk.Path & "\"
    
    If ReturnCode.ErrRT = ReadData(d) Then Exit Sub
    If ReturnCode.ErrRT = DataToStruct(d) Then Exit Sub
    
    If ReturnCode.ErrRT = OpenAllWk(d) Then Exit Sub
    
    If ReturnCode.ErrRT = GetResult(d) Then
        If VBA.MsgBox("�Ƿ�ر����й�������", vbYesNo) = vbNo Then Exit Sub
    End If
    
    If ReturnCode.ErrRT = CloseAllWk(d) Then Exit Sub
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "OK"
    
    Exit Sub
err_handle:
    MsgBox Err.Description
    If VBA.MsgBox("�Ƿ�ر����й�������", vbYesNo) = vbYes Then CloseAllWk d
End Sub

Private Function GetResult(d As DataStruct) As ReturnCode
    Dim i As Long
    
    For i = 0 To d.Count - 1
'        If i = 2 Then Stop
        
        Select Case d.uType(i)
        Case UpdateType.rng
            If ReturnCode.ErrRT = GetResultRange(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
            
        Case UpdateType.ColRelation
            If ReturnCode.ErrRT = GetResultColRelation(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
            
        Case UpdateType.ColRelationAppend
            If ReturnCode.ErrRT = GetResultColRelationAppend(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
            
        Case UpdateType.Formula
            If ReturnCode.ErrRT = GetResultFormula(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
        
        Case UpdateType.dic
            If ReturnCode.ErrRT = GetResultDic(d, i, False) Then
                GetResult = ErrRT
                Exit Function
            End If
            
        Case UpdateType.dicExists
            If ReturnCode.ErrRT = GetResultDic(d, i, True) Then
                GetResult = ErrRT
                Exit Function
            End If
        
        Case UpdateType.AddName
            If ReturnCode.ErrRT = GetResultAddName(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
        
        Case UpdateType.AddNo
            If ReturnCode.ErrRT = GetResultAddNo(d, i) Then
                GetResult = ErrRT
                Exit Function
            End If
             
        Case Else
            MsgBox "δ֪����" & VBA.CStr(d.uType(i))
            GetResult = ErrRT
            Exit Function
        End Select
    Next
    
    GetResult = SuccessRT
End Function

'������
Private Function GetResultAddNo(d As DataStruct, i As Long) As ReturnCode
    Dim shtDes As Worksheet
    Dim RowDes As Long
    
    'Ŀ�깤����
    Set shtDes = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName)
    shtDes.AutoFilterMode = False
    RowDes = shtDes.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row
    If RowDes < d.Des(i).RowStart Then
        d.sht.Cells(i, Pos.BeiZhu + 1).Value = "Des����û�г���RowStart"
        '����ʾ�������ش���
        GetResultAddNo = SuccessRT
        Exit Function
    End If
    
    With shtDes.Cells(d.Des(i).RowStart, d.Des(i).Action)
        .Value = 1
        .AutoFill Destination:=.Resize(RowDes - d.Des(i).RowStart + 1, 1), Type:=xlFillSeries
    End With
    
End Function


'����Զ�������
Private Function GetResultAddName(d As DataStruct, i As Long) As ReturnCode
    '��ȡ����Դ
    Dim shtSrc As Worksheet, RowSrc As Long
    Dim ArrKey(), ArrValue()
    
    Set shtSrc = d.ArrWk(d.dicWk(d.Src(i).wkName)).wk.Worksheets(d.Src(i).shtName)
    shtSrc.AutoFilterMode = False
    RowSrc = shtSrc.Cells(Cells.Rows.Count, d.Src(i).ColKey).End(xlUp).Row
    If RowSrc < d.Src(i).RowStart Then
        d.sht.Cells(i, Pos.BeiZhu + 1).Value = "Src����û�г���RowStart"
        '����ʾ�������ش���
        GetResultAddName = SuccessRT
        Exit Function
    Else
        ArrKey = shtSrc.Cells(1, d.Src(i).ColKey).Resize(RowSrc, 1).Value
        ArrValue = shtSrc.Cells(1, d.Src(i).Action).Resize(RowSrc, 1).Value
         
        If VBA.Len(d.Des(i).wkName) Then AddNameToWk ArrKey, ArrValue, d.ArrWk(d.dicWk(d.Des(i).wkName)).wk
        If VBA.Len(d.Des(i).shtName) Then AddNameToWk ArrKey, ArrValue, d.ArrWk(d.dicWk(d.Des(i).shtName)).wk
        If VBA.Len(d.Des(i).Action) Then AddNameToWk ArrKey, ArrValue, d.ArrWk(d.dicWk(d.Des(i).Action)).wk
    End If

    GetResultAddName = SuccessRT
End Function

Private Function AddNameToWk(ArrKey(), ArrValue(), wk As Workbook)
    Dim i As Long
    Dim strKey As String
    Dim strValue As String
    
    For i = 2 To UBound(ArrKey)
        strKey = VBA.CStr(ArrKey(i, 1))
        strValue = VBA.CStr(ArrValue(i, 1))
        If VBA.Len(strValue) = 0 Then strValue = "0"
        
        If VBA.Len(strKey) Then
            wk.Names.Add Name:=strKey, RefersToR1C1:="=" & strValue
        End If
    Next
End Function

Private Function GetResultDic(d As DataStruct, i As Long, bDicExists As Boolean) As ReturnCode
    'Դ��λ����key����Ԫ������item
    Dim shtDes As Worksheet
    Dim shtSrc As Worksheet
    Dim dic As Object, j As Long, strKey As String
    Dim ArrKey(), ArrItem()
    
    Set dic = CreateObject("Scripting.Dictionary")
    
    Dim RowDes As Long, RowSrc As Long
    '�ҵ�����Դ�ķ�Χ
    Set shtSrc = d.ArrWk(d.dicWk(d.Src(i).wkName)).wk.Worksheets(d.Src(i).shtName)
    shtSrc.AutoFilterMode = False
    RowSrc = shtSrc.Cells(Cells.Rows.Count, d.Src(i).ColKey).End(xlUp).Row
    If RowSrc < d.Src(i).RowStart Then
        d.sht.Cells(i, Pos.BeiZhu + 1).Value = "Src����û�г���RowStart"
        '����ʾ�������ش���
        GetResultDic = SuccessRT
        Exit Function
    Else
        'Ŀ�깤����
        Set shtDes = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName)
        shtDes.AutoFilterMode = False
        RowDes = shtDes.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row
        If RowDes < d.Des(i).RowStart Then
            d.sht.Cells(i, Pos.BeiZhu + 1).Value = "Des����û�г���RowStart"
            '����ʾ�������ش���
            GetResultDic = SuccessRT
            Exit Function
        End If
    
        ArrKey = shtSrc.Cells(1, d.Src(i).ColKey).Resize(RowSrc, 1).Value
        ArrItem = shtSrc.Cells(1, d.Src(i).Action).Resize(RowSrc, 1).Value
        For j = d.Src(i).RowStart To RowSrc
            dic(VBA.UCase$(VBA.CStr(ArrKey(j, 1)))) = ArrItem(j, 1)
        Next
        
        '���
        ArrKey = shtDes.Cells(1, d.Des(i).ColKey).Resize(RowDes, 1).Value
        ArrItem = shtDes.Cells(1, d.Des(i).Action).Resize(RowDes, 1).Value
        For j = d.Des(i).RowStart To RowDes
            strKey = VBA.CStr(ArrKey(j, 1))
            strKey = VBA.UCase$(strKey)
            If dic.Exists(strKey) Then
                ArrItem(j, 1) = dic(strKey)
            Else
                If Not bDicExists Then
                    'û�о����ԭ��������
                    ArrItem(j, 1) = ""
                End If
            End If
        Next
        
        shtDes.Cells(1, d.Des(i).Action).Resize(RowDes, 1).Value = ArrItem
    End If
    
    GetResultDic = SuccessRT
    
    Set dic = Nothing
    Set shtDes = Nothing
    Set shtSrc = Nothing
End Function

Private Function GetResultColRelation(d As DataStruct, i As Long) As ReturnCode
    '�����Ҫ�����һ��Des
    Dim shtDes As Worksheet
    Dim RowDes As Long
    
    Set shtDes = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName)
    shtDes.AutoFilterMode = False
    RowDes = shtDes.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row
    If RowDes >= d.Des(i).RowStart Then
        shtDes.Rows(VBA.CStr(d.Des(i).RowStart) & ":" & VBA.CStr(RowDes)).ClearContents
    End If
    
    GetResultColRelation = GetResultColRelationAppend(d, i)
End Function
Private Function GetResultColRelationAppend(d As DataStruct, i As Long) As ReturnCode
    Dim tmpSrc, tmpDes
    Dim shtDes As Worksheet
    Dim shtSrc As Worksheet
    
    tmpSrc = VBA.Split(d.Src(i).Action, SPLIT_WORD)
    tmpDes = VBA.Split(d.Des(i).Action, SPLIT_WORD)
    
    Dim iCount As Long
    iCount = UBound(tmpSrc) + 1
    If iCount <> UBound(tmpDes) + 1 Then
        MsgBox "��û��һһ��Ӧ��" & vbNewLine & "���������У�" & VBA.CStr(i)
        GetResultColRelationAppend = ErrRT
        Exit Function
    End If
    
    Dim j As Long
    Dim RowDes As Long, RowSrc As Long
    '�ҵ�����Դ�ķ�Χ
    Set shtSrc = d.ArrWk(d.dicWk(d.Src(i).wkName)).wk.Worksheets(d.Src(i).shtName)
    shtSrc.AutoFilterMode = False
    RowSrc = shtSrc.Cells(Cells.Rows.Count, d.Src(i).ColKey).End(xlUp).Row
    If RowSrc < d.Src(i).RowStart Then
        d.sht.Cells(i, Pos.BeiZhu + 1).Value = "Src����û�г���RowStart"
        '����ʾ�������ش���
        GetResultColRelationAppend = SuccessRT
        Exit Function
    Else
        '�ҵ�Ŀ�깤����������ʼ��
        Set shtDes = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName)
        shtDes.AutoFilterMode = False
        RowDes = shtDes.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row + 1
        If RowDes < d.Des(i).RowStart Then
            RowDes = d.Des(i).RowStart
        End If
        
        For j = 0 To iCount - 1
            shtDes.Cells(RowDes, VBA.CStr(tmpDes(j))).Resize(RowSrc - d.Src(i).RowStart + 1, 1).Value = shtSrc.Cells(d.Src(i).RowStart, VBA.CStr(tmpSrc(j))).Resize(RowSrc - d.Src(i).RowStart + 1, 1).Value
        Next
    End If
    
    '����һ�¸�ʽ
'    shtDes.Range(shtDes.Cells(d.Des(i).RowStart, 1), shtDes.Cells(1, 1).CurrentRegion.SpecialCells(xlCellTypeLastCell)).Borders.LineStyle = 1


    GetResultColRelationAppend = SuccessRT
End Function

Private Function GetResultFormula(d As DataStruct, i As Long) As ReturnCode
    '���d.Des(i).Action��¼�Ľ����кţ������RowStart��ColKey����λ
    '�ж��ұߵ�1���Ƿ�������
    Dim sht As Worksheet
    Dim rng As Range
    Dim i_row As Long
    
    Set sht = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName)
    If VBA.IsNumeric(VBA.Right$(d.Des(i).Action, 1)) Then
        Set rng = sht.Range(d.Des(i).Action)
    Else
        sht.AutoFilterMode = False
        i_row = sht.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row
        If i_row < d.Des(i).RowStart Then
            d.sht.Cells(i, Pos.BeiZhu + 1).Value = "����û�г���RowStart"
            '����ʾ�������ش���
            GetResultFormula = SuccessRT
            Exit Function
        Else
            Set rng = sht.Cells(d.Des(i).RowStart, d.Des(i).Action).Resize(i_row - d.Des(i).RowStart + 1, 1)
        End If
    End If
    
    If d.Src(i).Action = "=SUM" Then
        sht.AutoFilterMode = False
        i_row = sht.Cells(Cells.Rows.Count, d.Des(i).ColKey).End(xlUp).Row
        
        Set rng = sht.Range(d.Des(i).Action)
        'ͳ��rng���е����ĵ�Ԫ��
        rng.Formula = "=SUM(" & Cells(d.Des(i).RowStart, rng.Column).Resize(i_row - d.Des(i).RowStart + 1, 1).Address & ")"
    Else
        rng.FormulaR1C1Local = d.Src(i).Action
    End If
    
    GetResultFormula = SuccessRT
End Function

Private Function GetResultRange(d As DataStruct, i As Long) As ReturnCode
    '�����Ŀ�굥Ԫ��Χ��src��Ԫ��Χ�Ƿ�һ��
    '��һ�µ��������ʾһ��
    Dim RngSrc As Range, RngDes As Range
    Dim iRows1 As Long, iCols1 As Long
    Dim iRows2 As Long, iCols2 As Long
    
    On Error GoTo ErrHandle
    Set RngSrc = d.ArrWk(d.dicWk(d.Src(i).wkName)).wk.Worksheets(d.Src(i).shtName).Range(d.Src(i).Action)
    Set RngDes = d.ArrWk(d.dicWk(d.Des(i).wkName)).wk.Worksheets(d.Des(i).shtName).Range(d.Des(i).Action)
    
    iRows1 = RngSrc.Rows.Count
    iCols1 = RngSrc.Columns.Count
    
    iRows2 = RngDes.Rows.Count
    iCols2 = RngDes.Columns.Count
    
    If iRows1 <> iRows2 Or iCols1 <> iCols2 Then
        d.sht.Cells(i + Pos.RowStart, Pos.BeiZhu + 1).Value = "��Ԫ��Χ��һ��"
        Set RngDes = RngDes.Range("A1").Resize(iRows1, iCols1)
    End If
        
    RngDes.Value = RngSrc.Value
    
    GetResultRange = SuccessRT
    Exit Function
    
    '�п��ܵ�Ԫ���ַд����
ErrHandle:
    MsgBox Err.Description & vbNewLine & "���������У�" & VBA.CStr(i + Pos.RowStart)
    d.wk.Activate
    Cells(i, 1).Select
    
    GetResultRange = ErrRT
End Function

Private Function ReadData(d As DataStruct) As ReturnCode
    ActiveSheet.AutoFilterMode = False
    d.Rows = Cells(Cells.Rows.Count, Pos.DesWk).End(xlUp).Row
    If d.Rows < Pos.RowStart Then
        MsgBox "û������"
        ReadData = ReturnCode.ErrRT
        Exit Function
    End If
    d.Arr = Cells(1, 1).Resize(d.Rows, Pos.BeiZhu).Value
    '����±�ע�����һ�У���һ�л�������¼һЩ��ʾ��Ϣ������2����Ԫ��Χ��һ��
    d.sht.Cells(1, Pos.BeiZhu + 1).EntireColumn.Clear
    ReadData = ReturnCode.SuccessRT
End Function
'�����ݷŵ��ṹ����ȥ
Private Function DataToStruct(d As DataStruct) As ReturnCode
    Dim i As Long
    
    Dim dic As Object

    Set d.dicWk = CreateObject("Scripting.Dictionary")

    d.Count = d.Rows - Pos.RowStart + 1
    ReDim d.Src(d.Count - 1) As DataStructItem
    ReDim d.Des(d.Count - 1) As DataStructItem
    ReDim d.uType(d.Count - 1) As UpdateType
    Dim iTmp As Long
    
    For i = Pos.RowStart To d.Rows
        iTmp = i - Pos.RowStart
    
        d.Src(iTmp).wkName = VBA.CStr(d.Arr(i, Pos.SrcWk))
        d.Src(iTmp).shtName = VBA.CStr(d.Arr(i, Pos.SrcSht))
        d.Src(iTmp).RowStart = VBA.CLng(d.Arr(i, Pos.SrcRowStart))
        d.Src(iTmp).ColKey = VBA.CLng(d.Arr(i, Pos.SrcColKey))
        d.Src(iTmp).Action = VBA.CStr(d.Arr(i, Pos.SrcRng))
        
        d.Des(iTmp).wkName = VBA.CStr(d.Arr(i, Pos.DesWk))
        d.Des(iTmp).shtName = VBA.CStr(d.Arr(i, Pos.DesSht))
        d.Des(iTmp).RowStart = VBA.CLng(d.Arr(i, Pos.DesRowStart))
        d.Des(iTmp).ColKey = VBA.CLng(d.Arr(i, Pos.DesColKey))
        d.Des(iTmp).Action = VBA.CStr(d.Arr(i, Pos.DesRng))
        
        Select Case d.Arr(i, Pos.TheType)
        Case "Rng"
            d.uType(iTmp) = UpdateType.rng
        Case "ColRelation"
            d.uType(iTmp) = UpdateType.ColRelation
        Case "ColRelationAppend"
            d.uType(iTmp) = UpdateType.ColRelationAppend
        Case "Formula"
            d.uType(iTmp) = UpdateType.Formula
            If VBA.Left$(d.Src(iTmp).Action, 1) <> "=" Then d.Src(iTmp).Action = "=" & d.Src(iTmp).Action
        
        Case "dic"
            d.uType(iTmp) = UpdateType.dic
            
        Case "dicExists"
            d.uType(iTmp) = UpdateType.dicExists
        
        Case "AddName"
            d.uType(iTmp) = UpdateType.AddName
            'Ŀ�깤����������������¼�Ķ����ļ����ƣ�����Ҫ���
            RecordWk d.dicWk, d.Des(iTmp).shtName
            RecordWk d.dicWk, d.Des(iTmp).Action
            
        Case "AddNo"
            d.uType(iTmp) = UpdateType.AddNo
            
        Case Else
            MsgBox "δ֪����[" & Cells(i, Pos.TheType).Address(True, True) & "]"
            DataToStruct = ErrRT
            Exit Function
        End Select
        
        '��¼���й���������
        RecordWk d.dicWk, d.Src(iTmp).wkName
        
        If VBA.Len(d.Des(iTmp).wkName) = 0 Or VBA.Len(d.Des(iTmp).shtName) = 0 Or VBA.Len(d.Des(iTmp).Action) = 0 Then
            DataToStruct = ErrRT
            MsgBox "Ŀ�깤��������������Ԫ���ܶ�Ϊ�ա�" & "���������У�" & VBA.CStr(i)
            Exit Function
        End If
        
        RecordWk d.dicWk, d.Des(iTmp).wkName
    Next
      
    DataToStruct = SuccessRT
End Function
'��¼���й���������
Private Function RecordWk(dic As Object, wkName As String)
    If VBA.Len(wkName) Then
        If Not dic.Exists(wkName) Then dic(wkName) = dic.Count
    End If
End Function


Private Function OpenAllWk(d As DataStruct) As ReturnCode
    'Դ��������Ŀ�깤�������ܴ��ڽ������
    Dim i As Long
    Dim strKey As String
    ReDim d.ArrWk(d.dicWk.Count - 1) As WkType
       
    For i = Pos.RowStart To d.Rows
        strKey = VBA.CStr(d.Arr(i, Pos.SrcWk))
        'SrcWk  �п����ǿյģ���formula�����
        If VBA.Len(strKey) Then
            OpenWkItem d, strKey, False
        End If
        
        OpenWkItem d, VBA.CStr(d.Arr(i, Pos.DesWk)), True
        '��AddName״̬�£�Ŀ�깤����  Ŀ�굥Ԫ�� ��������д1���ļ�
        If d.Arr(i, Pos.TheType) = "AddName" Then
            OpenWkItem d, VBA.CStr(d.Arr(i, Pos.DesSht)), True
            OpenWkItem d, VBA.CStr(d.Arr(i, Pos.DesRng)), True
        End If
        
    Next i
    
    '��Ŀ�깤����������һ��Ҫ����bSave
    For i = Pos.RowStart To d.Rows
        strKey = VBA.CStr(d.Arr(i, Pos.SrcWk))
        
        d.ArrWk(OpenWkItem(d, VBA.CStr(d.Arr(i, Pos.DesWk)), True)).bSave = True
        '��AddName״̬�£�Ŀ�깤����  Ŀ�굥Ԫ�� ��������д1���ļ�
        If d.Arr(i, Pos.TheType) = "AddName" Then
            d.ArrWk(OpenWkItem(d, VBA.CStr(d.Arr(i, Pos.DesSht)), True)).bSave = True
            d.ArrWk(OpenWkItem(d, VBA.CStr(d.Arr(i, Pos.DesRng)), True)).bSave = True
        End If
        
    Next i
    
    OpenAllWk = SuccessRT
End Function

Private Function OpenWkItem(d As DataStruct, strKey As String, bSave As Boolean) As Long
    Dim pArr As Long
    
    pArr = d.dicWk(strKey)
    
    If d.ArrWk(pArr).wk Is Nothing Then
        If VBA.InStr(strKey, ":\") Then
            '���ǵ�ǰ�ļ����µ��ļ�
            Set d.ArrWk(pArr).wk = Workbooks.Open(strKey, False)
        Else
            '��û�д�
            Set d.ArrWk(pArr).wk = Workbooks.Open(d.Path & strKey, False)
            'Ŀ�깤���������Ҫ����
        End If
        d.ArrWk(pArr).wkName = strKey
        d.ArrWk(pArr).bSave = bSave
    End If
    
    OpenWkItem = pArr
End Function

Private Function CloseAllWk(d As DataStruct) As ReturnCode
    Dim i As Long
    
    For i = 0 To UBound(d.ArrWk)
        If Not d.ArrWk(i).wk Is Nothing Then
            d.ArrWk(i).wk.Close d.ArrWk(i).bSave
        End If
    Next
    
    CloseAllWk = SuccessRT
End Function
