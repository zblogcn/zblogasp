<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    haphic
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
Function CmtN_SendMessage(Byval strToAddress,Byval strToAddress2,Byval strReplyToAddress,Byval strFromName,Byval strSubject,Byval strBody)
	On Error Resume Next
	If Not InStr(strToAddress,"@")>0 And Not InStr(strToAddress2,"@")>0 Then
		CmtN_SendMessageViaJamil = False
		Application.Lock
		Application(ZC_BLOG_CLSID& "CmtN_LastMailLog")="您已关闭给站长发信的功能, 如果您确定您的SMTP服务器可以发送邮件, 那么没什么不正常的."
		Application.UnLock
		Exit Function
	End If
	CmtN_SendMessage=CmtN_SendMessageViaCDO(strToAddress,strToAddress2,strReplyToAddress,strFromName,strSubject,strBody)
	
End Function

Function CmtN_SendMessageViaCDO(Byval strToAddress,Byval strToAddress2,Byval strReplyToAddress,Byval strFromName,Byval strSubject,Byval strBody)
	Dim cdo
	Set cdo=Server.CreateObject("CDO.Message")
	IF Err.Number<>0 Then Exit Function
	With cdo.configuration.fields 
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")= CmtN_MailServerName 'SMTP 服务器地址 
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")= 25 '端口 25 
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")= CmtN_MailServerUserName '用户名 
		.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")= CmtN_MailServerUserPwd '用户密码 
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")= 1 'NONE, Basic (Base64 encoded), NTLM 
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout")= 10 '超时设置, 以秒为单位 
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False '是否使用套接字 true/false  
		.Update 
	End With
	cdo.To=strToAddress
	If strToAddress2<>"null" Then cdo.cc=strToAddress2
	cdo.From=CmtN_MailFromAddress
	'cdo.Sender=strFromName
	'CDO貌似没有定义发件人的方法。。
	cdo.Subject=strSubject
	cdo.HTMLBody=strBody
	cdo.HTMLBodyPart.Charset=CmtN_Charset
	cdo.Send
End Function
'*********************************************************
' 目的：    发送邮件
'*********************************************************
Function CmtN_SendMessageViaJm(Byval strToAddress,Byval strToAddress2,Byval strReplyToAddress,Byval strFromName,Byval strSubject,Byval strBody)

	'On Error Resume Next
	Dim jmail
	Set jmail = Server.CreateObject("JMAIL.Message") '建立发送邮件的对象
	If Err.Number<>0 Then CmtN_SendMessageViaJm=CmtN_SendMessageViaCDO(strToAddress,strToAddress2,strReplyToAddress,strFromName,strSubject,strBody):Exit Function
	
	jmail.Clear()

	jmail.Logging = True '记录发送日志
	jmail.silent = True '屏蔽例外错误，返回FALSE跟TRUE两值
	jmail.Charset = CmtN_Charset '邮件的文字编码为国标
	jmail.ContentType = "text/html" '邮件的格式为纯文本, 如text/html则为html格式

	Dim aryToAddress,ItemToAddress
	strToAddress=Replace(strToAddress,"，",",")
	If InStr(strToAddress,"@")>0 Then '管理员的邮件地址
		If InStr(strToAddress,",")>0 Then
			aryToAddress=Split(strToAddress,",")
			For Each ItemToAddress In aryToAddress
				If InStr(ItemToAddress,"@")>0 Then jmail.AddRecipient ItemToAddress '有多个管理员时
			Next
		Else
			jmail.AddRecipient strToAddress '只有一个管理员时
		End If
	End If
	If InStr(strToAddress2,"@")>0 Then jmail.AddRecipient strToAddress2 '第二收件人的地址
	'jmail.AddRecipient "haphic@126.com", "his name" '邮件收件人的地址, 后面的为可选项
	'jmail.AddRecipient "haphic@gmail.com" '邮件收件人的地址, 可重复加入

	jmail.From = CmtN_MailFromAddress '发件人的E-MAIL地址
	jmail.FromName = strFromName '发件人的姓名
	jmail.ReplyTo = strReplyToAddress '回复地址

	jmail.MailServerUserName = CmtN_MailServerUserName '登录邮件服务器所需的用户名
	jmail.MailServerPassword = CmtN_MailServerUserPwd '登录邮件服务器所需的密码

	JMail.Subject = strSubject '邮件的标题 
	JMail.HTMLBody = strBody

	jmail.Priority = 3'邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值

	If Not jmail.Send(CmtN_MailServerName) Then '执行邮件发送（通过邮件服务器地址）当要发送认证邮件时也可以使用格式：用户名:密码@邮件服务器
		If InStr(CmtN_MailServerAlternate,"@")>0 Then '启用备用发信服务器
			Dim h,f
			h=Replace(CmtN_MailServerAlternate," ","")
			h=Replace(h,"　","")
			h=Replace(h,"：",":")
			h=Replace(h,"（","(")
			h=Replace(h,"）",")")
			f=Mid(h,InStr(h,"(")+1,InStrRev(h,")")-InStr(h,"(")-1)
			h=Replace(h,"("&f&")","")
			jmail.From = f
			If jmail.Send(h) Then
				CmtN_SendMessageViaJamil = True
			Else
				CmtN_SendMessageViaJamil = False
			End If
		Else
			CmtN_SendMessageViaJamil = False
		End If
	Else
		CmtN_SendMessageViaJamil = True
	End If

	Application.Lock
	Application(ZC_BLOG_CLSID& "CmtN_LastMailLog")=jmail.log
	Application.UnLock

	jmail.Close() '关闭对象
	Set jmail=Nothing

Err.Clear
End Function

'*********************************************************
' 目的：    载入模板
'*********************************************************
Function CmtN_Template(ByVal strType)

	CmtN_Template=LoadFromFile(BlogPath &"zb_users/PLUGIN/CmtN/Templates/Template_"& strType &".html","utf-8")

End Function
'*********************************************************
%>