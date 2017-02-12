Attribute VB_Name = "模块1"
'代码由轩辕轼轲编写，如对Excel发邮件这个偏门主题感兴趣，可加入QQ群141616426，入群请输入验证代码：Excel发邮件
Sub CDOSENDEMAIL()
    Dim CDOMail As Variant
    
    On Error Resume Next                                         '出错后继续执行
    Application.DisplayAlerts = False                            '禁用系统提示
    ThisWorkbook.ChangeFileAccess Mode:=xlReadOnly               '将工作簿设置为只读模式
    
    Set CDOMail = CreateObject("CDO.Message")                    '创建对象
    CDOMail.From = "123@qq.com"                              '设置发信人的邮箱
    CDOMail.To = "321@qq.com"                                '设置收信人的邮箱
    CDOMail.Subject = "主题:用CDO发送邮件试验"                   '设定邮件的主题
    'CDOMail.TextBody = "文本内容"                               '使用文本格式发送邮件
    CDOMail.HtmlBody = "当您看到此封邮件，表明CDO设置正确"       '使用Html格式发送邮件
    CDOMail.AddAttachment ThisWorkbook.FullName                  '发送本工作簿为附件
    STUl = "http://schemas.microsoft.com/cdo/configuration/"     '微软服务器网址
    
    With CDOMail.Configuration.Fields
        .Item(STUl & "smtpserver") = "smtp.qq.com"               'SMTP服务器地址
        .Item(STUl & "smtpserverport") = 465                      'SMTP服务器端口
        .Item(STUl & "sendusing") = 2                            '发送端口
        .Item(STUl & "smtpauthenticate") = 1                     '远程服务器需要验证
        .Item(STUl & "sendusername") = "123"                 '发送方邮箱名称
        .Item(STUl & "sendpassword") = "111"                '发送方邮箱密码
        .Item(STUl & "smtpconnectiontimeout") = 60               '连接超时（秒）
        .Update
    End With
    
    CDOMail.Send                                                  '执行发送
    Set CDOMail = Nothing                                         '发送成功后即时释放对象
    If Err.Number = 0 Then
        MsgBox "成功发送邮件", , "温馨提示"                           '如果没有出错，则提示发送成功
    Else
        MsgBox Err.Description, vbInformation, "邮件发送失败"         '如果出错，则提示错误类型和错误代码
    End If
    
    ThisWorkbook.ChangeFileAccess Mode:=xlReadWrite               '将工作簿设置为读写模式
    Application.DisplayAlerts = True                              '恢复系统提示
End Sub

