Attribute VB_Name = "MTablesCompare"
Option Explicit

'���Ա�

Private Type CompareData
    Tables() As String 'ѡ��Աȵı��
    TablesAlias() As String '������������ǲ�ͬ�����ͬ�ֶ�����
    FieldsCondition() As String  '�����Աȵ������С���ͼ�š����Ƶ�
    FieldsData() As String      '�Ա�������� �������������۵�
    SQL As String
    SqlOther As String      '������Ҫ�˽�ģ������ȡ����ʱ����Ҫȥ���հ׵ȵ�
    rngOut As Range '�����Ԫ��
    
    wk As Workbook
    f As FCompare
End Type

Private cd As CompareData

Function Compare()
    If cd.f Is Nothing Or Not (cd.wk Is ActiveWorkbook) Then
        Set cd.f = New FCompare
        Set cd.wk = ActiveWorkbook
        cd.f.SheetNames = MFunc.GetSheetsName(cd.wk)
    End If
    'ͨ�����壬��ȡѡ��Ĺ������Ա��õ��ֶΡ�Ҫ�Ա�������ֶε���Ϣ
formShow:
    cd.f.Show
    If cd.f.bClose Then Set cd.f = Nothing: Exit Function
    
    If Not cd.f.Cancel Then
        cd.f.bSelect = False
        cd.Tables = cd.f.SheetNames()
        If Not cd.f.bSelect Then MsgBox "��ѡ��sheet": GoTo formShow: Exit Function
        
        cd.f.bSelect = False
        cd.FieldsCondition = cd.f.FieldsCondition()
        If Not cd.f.bSelect Then MsgBox "��ѡ�������Աȵ������С���ͼ�š����Ƶ�": GoTo formShow: Exit Function
        
        cd.f.bSelect = False
        cd.FieldsData = cd.f.FieldsData()
        If Not cd.f.bSelect Then MsgBox "��ѡ��Ա�������� �������������۵�": GoTo formShow: Exit Function
        
        
        GetAliasName cd '��ȡ���ı���
        AddBracesToTable cd '��������������
        
        cd.SqlOther = cd.f.SqlOther
        GetSql cd
        
        Set cd.rngOut = MFunc.GetRng("ѡ�������Ԫ��,������������������е�Ԫ��", "A1")
        If cd.rngOut Is Nothing Then Exit Function
        cd.rngOut.Parent.Cells.Clear
               
        ExcuteSql cd

    End If
End Function

Private Function GetSql(cd As CompareData) As Long
    '���Ȼ�ȡ���б��Ĳ��ظ���FieldsCondition
    Dim strFieldsCondition As String
    strFieldsCondition = VBA.Join(cd.FieldsCondition, ",")
    
    Dim strAll As String
    strAll = VBA.Join(cd.Tables, " " & cd.SqlOther & " Union Select " & strFieldsCondition & " From ")
    strAll = "(Select " & strFieldsCondition & " From " & strAll & " " & cd.SqlOther & ") A"
       
    Dim strAsc As String
    strAsc = VBA.Chr(VBA.Asc("A"))
    '���Left Join
    Dim i As Long, j As Long
    Dim strTmp As String, LeftJoinOn() As String
    ReDim LeftJoinOn(UBound(cd.FieldsCondition)) As String
    
    For i = 0 To UBound(cd.Tables)
        'Select A.F1,A.F����
        cd.SQL = "Select " & strAsc & ".*" ' & VBA.Join(cd.FieldsCondition, "," & strAsc & ".")
        
        'Table1.F1 As Table1F1,Table1.F2 As Table1F2,
        strTmp = ""
        For j = 0 To UBound(cd.FieldsData)
            strTmp = strTmp & "," & cd.Tables(i) & "." & cd.FieldsData(j) & " As " & cd.TablesAlias(i) & cd.FieldsData(j)
        Next j
        
        'Left Join xxTable On
        cd.SQL = cd.SQL & strTmp & " From " & strAll & " Left Join " & cd.Tables(i) & " On "
        
        'Table1.F1 = A.F1 And Table1.F2 = A.F2
        For j = 0 To UBound(cd.FieldsCondition)
            LeftJoinOn(j) = strAsc & "." & cd.FieldsCondition(j) & "=" & cd.Tables(i) & "." & cd.FieldsCondition(j)
        Next j
        
        cd.SQL = cd.SQL & VBA.Join(LeftJoinOn, " And ")
        
        strAsc = VBA.Chr(VBA.Asc(strAsc) + 1)
        If i < UBound(cd.Tables) Then strAll = "(" & cd.SQL & ") " & strAsc
    Next
    
    
End Function
'�������������� [xxx$]
Private Function AddBracesToTable(cd As CompareData) As Long
    Dim i As Long
    
    For i = 0 To UBound(cd.Tables)
        cd.Tables(i) = "[" & cd.Tables(i) & "$]"
    Next
End Function
'��ȡ��ı���
Private Function GetAliasName(cd As CompareData) As Long
    Dim i As Long, iLen As Long
    Dim j As Long
    Dim Dic As Object
    Dim k As Long
    'ÿ����ı�����������ͬ��������Ҫ�̣���������ȡ�ַ��������ֵ䣬�ж��Ƿ񶼲���ͬ
    
    Set Dic = CreateObject("Scripting.Dictionary")
    k = UBound(cd.Tables) + 1
    
    iLen = 99
    For i = 0 To k - 1
        If iLen > VBA.Len(cd.Tables(i)) Then
            iLen = VBA.Len(cd.Tables(i))
        End If
    Next
    
    ReDim cd.TablesAlias(k - 1) As String
    For i = 1 To iLen
        For j = 0 To k - 1
            cd.TablesAlias(j) = VBA.Left$(cd.Tables(j), i)
            Dic(cd.TablesAlias(j)) = 0
        Next j
        '�кͱ�һ�����Ƶĸ�����˵��������ͬ��
        If Dic.Count = k Then
            Exit Function
        Else
            Dic.RemoveAll
        End If
    Next
    'ѭ�����˻���û�еĻ���ֱ��ʹ��sheet����
    For j = 0 To k - 1
        cd.TablesAlias(j) = cd.Tables(j)
    Next j
End Function
Private Function ExcuteSql(cd As CompareData) As Long
    Dim c_ado As CADO

    Set c_ado = New CADO
    c_ado.SourceFile = cd.wk.FullName
    c_ado.SQL = cd.SQL
    c_ado.ResultToExcel cd.rngOut

    Set c_ado = Nothing
End Function
