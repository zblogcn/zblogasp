<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    haphic
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////

Class CmtN_Class
	Dim obj
	Dim ServerObject
	Dim objType
	Public mailTo
	Public mailTo2
	Public mailReply
	Public mailName
	Public mailSubject
	Public mailBody
	Public TemplateName
	Public ParentObj
	
	Private Ftemplate
	
	Public Property Let Template(strFileName)
		Ftemplate=LoadFromFile(BlogPath &"zb_users/PLUGIN/CmtN/Templates/Template_"& strFileName &".html","utf-8")
	End Property
	
	Public Property Get Template
		If Ftemplate<>"" Then
			Template = Ftemplate
			Exit Property
		Else
			Dim s
			s=LoadFromFile(BlogPath &"zb_users/PLUGIN/CmtN/Templates/Template_"& TemplateName &".html","utf-8")
			Ftemplate = s
			Template = Ftemplate
		End If
	End Property
	
	
	Function MakeCommentTemplate(objComment,isParent)

	
		Dim MT : MT = ""
		Dim MC : MC = FTemplate
		Dim Art_FirstName : Art_FirstName=""
		Dim Cmt_FirstName : Cmt_FirstName=""
		Dim Content
		Dim User
		Dim objRS
		Dim objArticle
		Set objArticle=New TArticle
		If objArticle.LoadInfoByID(objComment.log_id) Then
			If isParent=False Then
				Art_FirstName=Users(objArticle.AuthorID).FirstName
				Call GetUser
				Cmt_FirstName = objComment.Author
				If objComment.AuthorID>0 Then
					For Each User in Users
						If IsObject(User) Then
							If User.ID=objComment.AuthorID Then
								Cmt_FirstName = User.FirstName
								Exit For
							End If
						End If
					Next
				End If
				MC = Replace(MC,"<#Cmt_Article/author/name#>",Art_FirstName)
				If objComment.ParentID=0 Then
					MT = MT & Cmt_FirstName & " 在您的博客 """& ZC_BLOG_NAME &""" 里评论"
				Else
					MT = MT & Cmt_FirstName & "回复了您的评论"
				End If
				MC = Replace(MC,"<#Cmt_Type#>","评论")
				MC = Replace(MC,"<#Cmt_Article/title#>",objArticle.Title)
				MC = Replace(MC,"<#Cmt_Article/url#>",objArticle.Url)
				MC = Replace(MC,"<#Cmt_Article/PostTime#>",objArticle.PostTime)

				MC = Replace(MC,"<#Cmt_Article/author/name#>","")


				MC = Replace(MC,"<#Cmt_Url#>",objArticle.Url & "#cmt" & objComment.ID)
				
				MC = Replace(MC,"<#BLOG_LINK#>","<a href="""& BlogHost &""" title="""& ZC_BLOG_SUB_NAME &""" target=""_blank"">"& ZC_BLOG_NAME &"</a>")
			
				MC = Replace(MC,"<#Cmt_Author/name#>",Cmt_FirstName)
				MC = Replace(MC,"<#Cmt_Author/url#>",objComment.Homepage)
				MC = Replace(MC,"<#Cmt_Author/email#>",objComment.email)
				MC = Replace(MC,"<#Cmt_Author/IP#>",objComment.ip)
				MC = Replace(MC,"<#Cmt_Author/agent#>",objComment.agent)
			
				If Len(objComment.Homepage)> 0 Then
					MC = Replace(MC,"<#Cmt_Author#>","<a href="""& objComment.Homepage &""" target=""_blank"">"& Cmt_FirstName &"</a>")
				Else
					MC = Replace(MC,"<#Cmt_Author#>",Cmt_FirstName)
				End If
				MC = Replace(MC,"<#Cmt_PostTime#>",GetTime(Now()))
				
				If objComment.ParentID=0 Then
					Content=TContent(objComment.Content)
					MC = Replace(MC,"<#Cmt_Content#>",Content)
					MC = Replace(MC,"<#MAIL_RECEIVER#>",Art_FirstName)
				Else
					Content=TContent(objComment.Content)
					MC = Replace(MC,"<#Cmt_Content_Child#>",Content)
					If IsObject(ParentObj)=False Then GetParentObj(objComment.ParentID)
					MC = Replace(MC,"<#Cmt_Content#>",TContent(ParentObj.Content))
					Cmt_FirstName = ParentObj.Author
					For Each User in Users
						If IsObject(User) Then
							If User.ID=ParentObj.AuthorID Then
								Cmt_FirstName = User.FirstName
								Exit For 
							End If
						End If
					Next
					MC = Replace(MC,"<#MAIL_RECEIVER#>",Cmt_FirstName)

				End If
				
				
			End If
		End If
		mailSubject=MT
		mailBody=MC
		MakeCommentTemplate=MC


	End Function
	
	Function TContent(str)
		TContent=TransferHTML(UBBCode(str,"[link][link-antispam][font][face]"),"[enter][nofollow]")
	End Function
	Function GetParentObj(ParentID)
		Dim o
		Set o=New TComment
		o.LoadInfoById ParentID
		Set GetParentObj=o
		Set ParentObj=o
	End Function
	
	Function InitObj
		On Error Resume Next
		Set obj=CreateObject(ServerObject)
		If Err.Number=0 Then InitObj=True
	End Function
	
	Sub Class_Initialize
		On Error Resume Next
		ServerObject="JMail.Message"
		objType=1
		If InitObj=False Then
			ServerObject="CDO.Message"
			InitObj
			objType=2
			With obj.configuration.fields 
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
		Else
			obj.Clear()
			obj.Logging = True '记录发送日志
			obj.silent = True '屏蔽例外错误，返回FALSE跟TRUE两值
			obj.Charset = CmtN_Charset '邮件的文字编码为国标
			obj.ContentType = "text/html" '邮件的格式为纯文本, 如text/html则为html格式
			obj.MailServerUserName = CmtN_MailServerUserName '登录邮件服务器所需的用户名
			obj.MailServerPassword = CmtN_MailServerUserPwd '登录邮件服务器所需的密码
		End If
		MailTo="null"
		MailTo2="null"
	End Sub
	
	Function Send()
		On Error Resume Next
		If MailTo="null" And MailTo2="null" Then Exit Function
		If MailTo="null" And MailTo2<>"null" Then Mailto=Mailto2:Mailto2="null" 
		If Not InStr(MailTo,"@")>0 And Not InStr(MailTo2,"@")>0 Then
			Send = False
			Application.Lock
			Application(ZC_BLOG_CLSID& "CmtN_LastMailLog")="您已关闭给站长发信的功能, 如果您确定您的SMTP服务器可以发送邮件, 那么没什么不正常的."
			Application.UnLock
			Exit Function
		End If
		Dim aryToAddress,ItemToAddress
		Select Case objType
			Case 2
				obj.To=Split(MailTo,",")(0)
				If MailTo2<>"null" Then obj.cc=MailTo2
				obj.From=CmtN_MailFromAddress
				'cdo.Sender=mailName
				'CDO貌似没有定义发件人的方法。。
				obj.Subject=mailSubject
				obj.HTMLBody=mailBody
				obj.HTMLBodyPart.Charset=CmtN_Charset
				obj.Send
				Send=True
			Case 1
				MailTo=Replace(MailTo,"，",",")
				If InStr(MailTo,"@")>0 Then '管理员的邮件地址
					If InStr(MailTo,",")>0 Then
						aryToAddress=Split(MailTo,",")
						For Each ItemToAddress In aryToAddress
							If InStr(ItemToAddress,"@")>0 Then obj.AddRecipient ItemToAddress '有多个管理员时
						Next
					Else
						obj.AddRecipient MailTo '只有一个管理员时
					End If
				End If
				If InStr(MailTo2,"@")>0 Then obj.AddRecipient MailTo2 '第二收件人的地址
				'jmail.AddRecipient "haphic@126.com", "his name" '邮件收件人的地址, 后面的为可选项
				'jmail.AddRecipient "haphic@gmail.com" '邮件收件人的地址, 可重复加入
			
				obj.From = CmtN_MailFromAddress '发件人的E-MAIL地址
				obj.FromName = mailName '发件人的姓名
				obj.ReplyTo = mailReply '回复地址
				obj.Subject = mailSubject '邮件的标题 
				obj.HTMLBody = mailBody
				obj.Priority = 3'邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
			
				If Not obj.Send(CmtN_MailServerName) Then '执行邮件发送（通过邮件服务器地址）当要发送认证邮件时也可以使用格式：用户名:密码@邮件服务器
					If InStr(CmtN_MailServerAlternate,"@")>0 Then '启用备用发信服务器
						Dim h,f
						h=Replace(CmtN_MailServerAlternate," ","")
						h=Replace(h,"　","")
						h=Replace(h,"：",":")
						h=Replace(h,"（","(")
						h=Replace(h,"）",")")
						f=Mid(h,InStr(h,"(")+1,InStrRev(h,")")-InStr(h,"(")-1)
						h=Replace(h,"("&f&")","")
						obj.From = f
						If obj.Send(h) Then
							Send = True
						Else
							Send = False
						End If
					Else
						Send = False
					End If
				Else
					Send = True
				End If
			
				Application.Lock
				Application(ZC_BLOG_CLSID& "CmtN_LastMailLog")=obj.log
				Application.UnLock
			
				obj.Close() '关闭对象
		End Select
	End Function
	
	Function LoadTemplate(name)
	End Function
	
End Class

'*********************************************************
' 目的：    发送邮件
'*********************************************************
Function CmtN_SendMessage(Byval MailTo,Byval MailTo2,Byval mailReply,Byval mailName,Byval mailSubject,Byval mailBody)
	CmtN.mailTo=MailTo
	CmtN.mailTo2=MailTo2
	cmtN.mailReply=mailReply
	cmtn.mailname=mailName
	cmtn.mailsubject=mailSubject
	cmtn.mailBody=mailBody
	CmtN_SendMessage=cmtn.send
End Function

Function CMTN_SendOutGoingMails
	CmtN_Initialize
	On Error Resume Next

	If CMTN_MailSendDelay = False Then Exit Function

	'Application.Contents.RemoveAll    'Test Olny

	Dim strOperateTime

	Application.Lock
	strOperateTime=Application(ZC_BLOG_CLSID & "CMTN_OPERATETIME")
	Application.UnLock

	If strOperateTime=Empty Then strOperateTime=DateAdd("d",-60,Now())

	If DateDiff("s",strOperateTime,Now())>CMTN_MailSendDelayTime Then

		Dim aryFileList
		aryFileList=LoadIncludeFiles("zb_users/PLUGIN/CMTN/OutGoingMails/")

		Dim fso
		Set fso = Server.CreateObject("Scripting.FileSystemObject")

		Dim l,c
		Dim m,n
		For Each l In aryFileList
			If Not IsEmpty(l) Then
				If (Not LCase(l)="index.html") Then

					c=LoadFromFile(BlogPath & "zb_users/PLUGIN/CMTN/OutGoingMails/"& l,"utf-8")

					If len(c)>3 Then
						m=Split(c,VbCrlf)
						n=Replace(c,m(0)&vbCrlf&m(1)&vbCrlf&m(2)&vbCrlf,"")
						Call CmtN_SendMessage(m(1),m(2),CMTN_MailReplyToAddress,CMTN_MailFromName,m(0),n)
					End If

					fso.DeleteFile(BlogPath & "zb_users/PLUGIN/CMTN/OutGoingMails/"& l)
					Exit For

				End If
			End If
		Next

		Set fso = Nothing

		Application.Lock
		Application(ZC_BLOG_CLSID & "CMTN_OPERATETIME")=Now()
		Application.UnLock

	End if

Err.Clear
End Function
%>