Attribute VB_Name = "MAnswer"
Option Explicit

'��ȡ��
'�𰸵�word��1���Ծ�1�����
'   ��1�������
'   ��2���������
'   ��3���ǵ�ѡ
'   ��4���Ƕ�ѡ
'   ��5���Ǽ������Ŀ���Ѿ��д𰸣�����Ҫ��
Function GetAnswer(d As DataStruct)
    Dim i As Long
    Dim doc As Document
    Dim preType As Long
    Dim pIndex As Long
    Dim j As Long
    
    Set doc = Documents.Open(ThisDocument.Path & "\��.docx")
    
    preType = -1
    pIndex = -1
    For i = 1 To d.Next - 1
        '����Ѿ����˴𰸵�
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
