'http://club.excelhome.net/thread-1256518-3-1.html
'���û��schema.ini��ADOĬ���ı��ļ����ö��ŷָ��ģ������schema.ini��ADO�Ͱ�����ָ���ķָ������Ƿ��б����еȽ��в�ѯ

Function SaveTxtToAccess(sFIELD As String, txt_name As String) As Long
    Dim str_sql As String
    Dim s As String
    Dim mypath
    
    mypath = ThisWorkbook.Path & "\"

    s = "[" & txt_name & "]" & vbCrLf & "COLNAMEHEADER = TRUE" & vbCrLf & "Format = Delimited" & vbCrLf & "Col1=��Ʊ���� Char 6" & vbCrLf & "Col2=��Ʊ���� Char" & _
            vbCrLf & "Col3=���� DATE" & vbCrLf & "Col4=���� Double" & vbCrLf & "Col5=��� Double" & vbCrLf & "Col6=��� Double" & vbCrLf & "Col7=���� Double" _
             & vbCrLf & "Col8=�ǵ��� Double" & vbCrLf & "Col9=�ɽ��� Double" & vbCrLf & "Col10=�ɽ��� Double"

    Open mypath & "schema.ini" For Output As #1
    Print #1, s
    Close #1
    Kill mypath & "schema.ini"
        
    str_sql = "Insert Into [" & sTABLE_NAME & "] (" & sFIELD & ") " & _
                "Select * From [Text;Database=" & ThisWorkbook.Path & ";].[" & txt_name & "]"

    ExecuteSql str_sql
End Function