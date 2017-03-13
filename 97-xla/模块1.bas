Attribute VB_Name = "模块1"
Option Explicit


Sub changYongMingLing()
    Dim usf As Object
    Set usf = New UserForm1
    usf.Show 0
    Set usf = Nothing
End Sub


Sub usfSqlShow()
    usfSql.Show 0
End Sub


Sub usfFileShow()
    usfFile.Show 0
End Sub

Sub getShtName()
    Dim Arr() As String, wk As Workbook
    Dim rng As Range
    
    Func.getRngByInputBox rng
    If rng Is Nothing Then GoTo A
            
    Set wk = ActiveWorkbook
    Func.getShtNameFromWorkbook wk, Arr, False
    
    Arr(0) = "SheetName"
    rng.Resize(UBound(Arr) + 1, 1).value = Application.WorksheetFunction.Transpose(Arr)
A:
    Erase Arr
    Set rng = Nothing
End Sub

Sub isAddin()
    ThisWorkbook.isAddin = Not ThisWorkbook.isAddin
End Sub

Sub fanYi()
    Dim rng As Range
    Dim cm As Comment
    Dim str As String
    
    If VBA.TypeName(Selection) = "Range" Then
        For Each rng In Selection
            If rng.value <> "" Then
                If rng.Comment Is Nothing Then
                    rng.AddComment tran_cell(rng)
                Else
                    Set cm = rng.Comment
                    cm.Text cm.Text & vbNewLine & tran_cell(rng)
                End If
            End If
        Next
    End If
    
    Set rng = Nothing
    Set cm = Nothing
End Sub

Sub fanYi_2()
    Dim rng As Range
    Dim str As String
    
    If VBA.TypeName(Selection) = "Range" Then
        For Each rng In Selection
            If rng.value <> "" Then
                rng.value = rng.value & vbNewLine & tran_cell(rng)
            End If
        Next
    End If
    
    Set rng = Nothing
End Sub

Function tran_cell(rng As Range) As String
    Dim str_html As String
    Dim xml As Object
    Dim regx As Object
    Set regx = CreateObject("VBScript.Regexp")
    
    Set xml = CreateObject("Microsoft.XMLHTTP")
    
    str_html = translate(xml, rng.value)
    
    If str_html = "error" Then
        tran_cell = str_html
    Else
        tran_cell = json(str_html)
    End If

    Set xml = Nothing
    Set regx = Nothing
End Function

Function translate(xml As Object, str_word As String) As String
    On Error GoTo ErrHandle
    
    With xml
        .Open "POST", "http://fanyi.youdao.com/translate", False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send "i=" & str_word & "&doctype=json"
        translate = .responsetext
'        translate = Split(translate, """tgt"":""")(1)
'        translate = Split(translate, """}]]}")(0)
    End With
    
'    StrConv(.ResponseBody, vbUnicode)
    Exit Function
    
ErrHandle:
    translate = "error"
End Function

Function json(str_html As String) As String
    Dim objJSON As Object
    Dim Cell '这里不能定义为object类型
    Dim tmp
    
    On Error GoTo ErrHandle
    
    With CreateObject("msscriptcontrol.scriptcontrol")
        .Language = "JavaScript"
        .AddCode "var mydata =" & str_html
        Set objJSON = .CodeObject
    End With
'    Stop '查看vba本地窗口里objJSON对象以了解JSON数据在vba里的形态
    For Each Cell In objJSON.mydata.translateResult
        For Each tmp In Cell
            json = json & tmp.tgt
        Next tmp
    Next
    
Exit Function
    
ErrHandle:
    json = "error"
End Function

Sub DeleteLeftTen()

    ActiveSheet.AutoFilterMode = False

    Rows("10:" & Cells.Rows.Count).Delete
End Sub
