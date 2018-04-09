Attribute VB_Name = "MMain"
Option Explicit

Enum Question
    No  '试卷
    Answer
    Text
    Type_ '题型
    Score '分数
    
    Count
End Enum

Enum QuestionType
    TianKong
    PanDuan
    DanXuan
    DuoXuan
    JianDa
End Enum

Type DataStruct
    Rows As Long
    No As Long
    Src As Variant
    Result() As Variant
    Next As Long
    QuestionType As QuestionType
End Type

Sub VBAMain()
    Dim i As Long
    Dim str As String, tmp As String
    Dim d As DataStruct
    
    str = ThisDocument.Content.Text
    d.Src = VBA.Split(str, VBA.Chr(13))
        
    d.Rows = UBound(d.Src)
    i = 0
    ReDim d.Result(d.Rows, Question.Count - 1) As Variant
    d.Result(0, Question.No) = "试卷"
    d.Result(0, Question.Type_) = "类型"
    d.Result(0, Question.Text) = "题目"
    d.Result(0, Question.Answer) = "答案"
    d.Result(0, Question.Score) = "分数"
    
    d.Next = 1
    Do Until i > d.Rows
        str = VBA.CStr(d.Src(i))
        
        'If VBA.InStr(str, "不符合涉密会议管理要求的做法是") Then Stop
        
        If VBA.Len(str) = 0 Then
            
        ElseIf VBA.InStr(str, "洪都公司涉密人员保密知识应知应会测试试题") Then
            d.No = d.No + 1
            
        ElseIf IsNewTypeBegin(str, d.QuestionType) Then
            
        Else
           
            If GetText(d, i, str) = -1 Then Exit Sub
        End If
        
        i = i + 1
    Loop
    
    GetAnswer d
    OutData d
End Sub

Private Function IsNewTypeBegin(str As String, ByRef QuestionType As QuestionType) As Boolean
    IsNewTypeBegin = True
    If VBA.Mid$(str, 3, 3) = "填空题" Then
        QuestionType = TianKong
    ElseIf VBA.Mid$(str, 3, 3) = "判断题" Then
        QuestionType = PanDuan
    ElseIf VBA.Mid$(str, 3, 3) = "单项选" Then
        QuestionType = DanXuan
    ElseIf VBA.Mid$(str, 3, 3) = "多项选" Then
        QuestionType = DuoXuan
    ElseIf VBA.Mid$(str, 3, 3) = "简答题" Then
        QuestionType = JianDa
    ElseIf VBA.Mid$(str, 3, 3) = "论述题" Then
        QuestionType = JianDa
        
    Else
        IsNewTypeBegin = False
    End If
End Function

Private Function GetText(d As DataStruct, ByRef Index As Long, StrText As String) As Long
    Dim tmp As String
    Dim strAns As String
    
    If d.QuestionType = TianKong Or d.QuestionType = PanDuan Then
        '单选和判断都是一行的
    Else
        
        Index = Index + 1
        tmp = VBA.CStr(d.Src(Index))
        tmp = VBA.Trim$(tmp)
         '找到下一行以数字开头的
        Do Until VBA.IsNumeric(VBA.Left$(tmp, 1)) Or IsNewTypeBegin(tmp, 1) Or (VBA.InStr(tmp, "洪都公司涉密人员保密知识应知应会测试试题")) > 0
            If VBA.Len(tmp) Then
                tmp = VBA.Trim$(tmp)
                
                If d.QuestionType = DanXuan Or d.QuestionType = DuoXuan Then
                    '单选和多选，要处理一下ABCDEF……分别在单独的一行
                    If SplitByAtoG(tmp) = -1 Then
                        GetText = -1
                        Exit Function
                    End If
                    StrText = StrText & tmp
                    
                Else
                    If VBA.Left$(tmp, 2) = "答：" Then tmp = VBA.Mid$(tmp, 3)
                    '简答题--就是答案了
                    If VBA.Len(strAns) Then
                        strAns = strAns & vbNewLine & tmp
                    Else
                        strAns = tmp
                    End If
                End If
                
            End If
            
            Index = Index + 1
            If Index > d.Rows Then Exit Do
            tmp = VBA.CStr(d.Src(Index))
        Loop
        Index = Index - 1
        
    End If
    
    d.Result(d.Next, Question.No) = d.No
    d.Result(d.Next, Question.Text) = StrText
    d.Result(d.Next, Question.Type_) = d.QuestionType
    
    If d.QuestionType = TianKong Then
        d.Result(d.Next, Question.Score) = 2
    ElseIf d.QuestionType = PanDuan Then
        d.Result(d.Next, Question.Score) = 1
    ElseIf d.QuestionType = DanXuan Then
        d.Result(d.Next, Question.Score) = 2
    ElseIf d.QuestionType = DuoXuan Then
        d.Result(d.Next, Question.Score) = 4
    Else
        d.Result(d.Next, Question.Score) = 10
    End If
    
    d.Result(d.Next, Question.Answer) = strAns
    d.Next = d.Next + 1
End Function

'将字符串分开
Private Function SplitByAtoG(ByRef str As String) As Long
    Dim AscA As Integer, AscG As Integer
    Dim AscTmp As Integer
    Dim i As Long, iPre As Long
    Dim strTmp As String
    
    AscA = VBA.Asc("A")
    AscG = VBA.Asc("G")
    
    AscTmp = VBA.Asc(VBA.Left$(str, 1))
    
    If AscTmp < AscA Or AscTmp > AscG Then
        MsgBox "首字符不是A-G。" & vbNewLine & str
        SplitByAtoG = -1
    Else
        iPre = 1
        i = 2
        Do Until i > VBA.Len(str)
            AscTmp = VBA.Asc(VBA.Mid$(str, i, 1))
            If AscTmp >= AscA And AscTmp <= AscG Then
                strTmp = strTmp & vbNewLine & VBA.Mid$(str, iPre, i - iPre)
                
                iPre = i
            End If
            i = i + 1
        Loop
        
        If VBA.Len(strTmp) Then
            strTmp = strTmp & vbNewLine & VBA.Mid$(str, iPre, i - iPre)
            str = strTmp
        Else
            str = vbNewLine & str
        End If
        
        SplitByAtoG = 1
    End If
End Function

Private Function OutData(d As DataStruct)
    Dim excel As Object
    
    Set excel = VBA.CreateObject("Excel.Application")
    
    excel.Visible = True
    
    Dim wk As Object
    Set wk = excel.Workbooks.Add
    wk.Worksheets(1).Cells(1, 1).Resize(d.Next, Question.Count).Value = d.Result
    wk.SaveAs ThisDocument.Path & "\题库"
End Function

