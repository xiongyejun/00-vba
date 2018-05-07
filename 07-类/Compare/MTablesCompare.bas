Attribute VB_Name = "MTablesCompare"
Option Explicit

'多表对比

Private Type CompareData
    Tables() As String '选择对比的表格
    TablesAlias() As String '别名，用来标记不同表的相同字段名称
    FieldsCondition() As String  '用来对比的条件列——图号、名称等
    FieldsData() As String      '对比输出的列 ——数量、单价等
    SQL As String
    SqlOther As String      '其他需要了解的，比如获取表格的时候，需要去除空白等等
    rngOut As Range '输出单元格
    
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
    '通过窗体，获取选择的工作表、对比用的字段、要对比输出的字段等信息
formShow:
    cd.f.Show
    If cd.f.bClose Then Set cd.f = Nothing: Exit Function
    
    If Not cd.f.Cancel Then
        cd.f.bSelect = False
        cd.Tables = cd.f.SheetNames()
        If Not cd.f.bSelect Then MsgBox "请选择sheet": GoTo formShow: Exit Function
        
        cd.f.bSelect = False
        cd.FieldsCondition = cd.f.FieldsCondition()
        If Not cd.f.bSelect Then MsgBox "请选择用来对比的条件列——图号、名称等": GoTo formShow: Exit Function
        
        cd.f.bSelect = False
        cd.FieldsData = cd.f.FieldsData()
        If Not cd.f.bSelect Then MsgBox "请选择对比输出的列 ——数量、单价等": GoTo formShow: Exit Function
        
        
        GetAliasName cd '获取表格的别名
        AddBracesToTable cd '给表格加上中括号
        
        cd.SqlOther = cd.f.SqlOther
        GetSql cd
        
        Set cd.rngOut = MFunc.GetRng("选择输出单元格,会清除输出工作表的所有单元格。", "A1")
        If cd.rngOut Is Nothing Then Exit Function
        cd.rngOut.Parent.Cells.Clear
               
        ExcuteSql cd

    End If
End Function

Private Function GetSql(cd As CompareData) As Long
    '首先获取所有表格的不重复的FieldsCondition
    Dim strFieldsCondition As String
    strFieldsCondition = VBA.Join(cd.FieldsCondition, ",")
    
    Dim strAll As String
    strAll = VBA.Join(cd.Tables, " " & cd.SqlOther & " Union Select " & strFieldsCondition & " From ")
    strAll = "(Select " & strFieldsCondition & " From " & strAll & " " & cd.SqlOther & ") A"
       
    Dim strAsc As String
    strAsc = VBA.Chr(VBA.Asc("A"))
    '多个Left Join
    Dim i As Long, j As Long
    Dim strTmp As String, LeftJoinOn() As String
    ReDim LeftJoinOn(UBound(cd.FieldsCondition)) As String
    
    For i = 0 To UBound(cd.Tables)
        'Select A.F1,A.F……
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
'给表格加上中括号 [xxx$]
Private Function AddBracesToTable(cd As CompareData) As Long
    Dim i As Long
    
    For i = 0 To UBound(cd.Tables)
        cd.Tables(i) = "[" & cd.Tables(i) & "$]"
    Next
End Function
'获取表的别名
Private Function GetAliasName(cd As CompareData) As Long
    Dim i As Long, iLen As Long
    Dim j As Long
    Dim Dic As Object
    Dim k As Long
    '每个表的别名都不能相同，尽量又要短，从左边逐个取字符，放入字典，判断是否都不相同
    
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
        '有和表一样名称的个数就说明都不相同了
        If Dic.Count = k Then
            Exit Function
        Else
            Dic.RemoveAll
        End If
    Next
    '循环完了还是没有的话就直接使用sheet名称
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
