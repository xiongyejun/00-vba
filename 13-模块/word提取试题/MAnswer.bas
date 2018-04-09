Attribute VB_Name = "MAnswer"
Option Explicit

'获取答案
'答案的word是1个试卷1个表格
'   第1列是序号
'   第2列是填空题
'   第3列是单选
'   第4列是多选
'   第5列是简答，在题目里已经有答案，不需要了
Function GetAnswer(d As DataStruct)
    Dim i As Long
    Dim doc As Document
    Dim preType As Long
    Dim pIndex As Long
    Dim j As Long
    
    Set doc = Documents.Open(ThisDocument.Path & "\答案.docx")
    
    preType = -1
    pIndex = -1
    For i = 1 To d.Next - 1
        '简答已经有了答案的
        If QuestionType.JianDa <> d.Result(i, Question.Type_) Then
            If d.Result(i, Question.Type_) <> preType Then
                preType = d.Result(i, Question.Type_)
                pIndex = 2
            Else
                pIndex = pIndex + 1
            End If
            
            d.Result(i, Question.Answer) = doc.Tables(d.Result(i, Question.No)).Cell(pIndex, d.Result(i, Question.Type_) + 2).Range.Text
            d.Result(i, Question.Answer) = VBA.Left$(d.Result(i, Question.Answer), VBA.Len(d.Result(i, Question.Answer)) - 2)
            
'            If d.Result(i, Question.Type_) = QuestionType.TianKong Then
'                d.Result(i, Question.Answer) = VBA.Replace(d.Result(i, Question.Answer), " ", "")
'            End If
        End If
    Next
    
    doc.Close False
End Function
