<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    haphic
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////

'*********************************************************
' 目的：    发送邮件
'*********************************************************
Function CmtN_SendMessageViaJamil(Byval strToAddress,Byval strToAddress2,Byval strReplyToAddress,Byval strFromName,Byval strSubject,Byval strBody)
On Error Resume Next

	If Not InStr(strToAddress,"@")>0 And Not InStr(strToAddress2,"@")>0 Then
		CmtN_SendMessageViaJamil = False
		Application.Lock
		Application(ZC_BLOG_CLSID& "CmtN_LastMailLog")="您已关闭给站长发信的功能, 如果您确定您的SMTP服务器可以发送邮件, 那么没什么不正常的."
		Application.UnLock
		Exit Function
	End If

	If CmtN_UseMailBrige Then

		Dim strServerFeedback,strServerURL
		Dim strSendingLog
		strServerURL=CmtN_MailBrigeDomain
		'If Right(strServerURL,1)<>"/" Then strServerURL=strServerURL & "/"
		strServerURL=strServerURL '&"zb_users/PLUGIN/CmtN/mailbridge.asp"

		Dim strSendData
		strSendData = "inpTo="& Server.URLEncode(strToAddress) &"&inpTo2="& Server.URLEncode(strToAddress2) &"&inpReplyTo="& Server.URLEncode(strReplyToAddress) &"&inpFrom="& Server.URLEncode(strFromName) &"&inpSubject="& Server.URLEncode(strSubject) &"&inpBody="& Server.URLEncode(strBody) &"&inpKey="& Server.URLEncode(CmtN_MailBrigeKey)

		Dim objPing
		Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP")
			objPing.SetTimeOuts 5000,5000,4000,10000

			objPing.open "POST",strServerURL,False

			objPing.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objPing.send strSendData

			If objPing.Status=200 Then
				strServerFeedback=objPing.ResponseText
			Else
				strServerFeedback = False
			End If
		Set objPing = Nothing


		If strServerFeedback=False Then
			CmtN_SendMessageViaJamil = False
			strSendingLog = "远程服务器连接失败!"
		ElseIf InStr(strServerFeedback,"[CONN-OK]")>0 Then

			If Cbool(Mid(strServerFeedback,11,1)) = False Then
				CmtN_SendMessageViaJamil = False
				strSendingLog = strServerFeedback
			Else
				CmtN_SendMessageViaJamil = True
				strSendingLog = strServerFeedback
			End If

		Else

			CmtN_SendMessageViaJamil = False
			strSendingLog = "远程服务器地址错误!"

		End If

		Application.Lock
		Application(ZC_BLOG_CLSID& "CmtN_LastMailLog")=strSendingLog
		Application.UnLock

	Exit Function
	End If


	Dim jmail
	Set jmail = Server.CreateObject("JMAIL.Message") '建立发送邮件的对象
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