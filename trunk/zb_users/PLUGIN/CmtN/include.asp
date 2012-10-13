<!-- #include file="function.asp" -->
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    haphic
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
Dim CmtN_Config
Dim CmtN_Charset 

Dim CmtN_MailToAddress 
Dim CmtN_MailReplyToAddress 
Dim CmtN_MailFromName 
Dim CmtN_NotifyCmtLeaver 

Dim CmtN_MailServerName 
Dim CmtN_MailServerUserName 
Dim CmtN_MailServerUserPwd 
Dim CmtN_MailFromAddress 

Dim CmtN_MailServerAlternate 

Dim CmtN_MailSendDelay 
Dim CmtN_MailSendDelayTime 

Dim CmtN_UseMailBrige 
Dim CmtN_MailBrigeDomain 
Dim CmtN_MailBrigeKey 
Dim CmtN_MailBrigeForOthers 


'注册插件
Call RegisterPlugin("CmtN","ActivePlugin_CmtN")


'具体的接口挂接
Function ActivePlugin_CmtN() 

	'挂上接口
	'Action_Plugin_CommentPost_Succeed
	Call Add_Filter_Plugin("Filter_Plugin_CommentPost_Succeed","Filter_Plugin_PostComment_Succeed")

	'挂上接口
	'Action_Plugin_Catalog_End
	Call Add_Action_Plugin("Action_Plugin_Catalog_End","CmtN_SendOutGoingMails")
	'Action_Plugin_Default_End
	Call Add_Action_Plugin("Action_Plugin_Default_End","CmtN_SendOutGoingMails")
	'Action_Plugin_Tags_End
	Call Add_Action_Plugin("Action_Plugin_Tags_End","CmtN_SendOutGoingMails")
	'Action_Plugin_View_End
	Call Add_Action_Plugin("Action_Plugin_View_End","CmtN_SendOutGoingMails")


End Function



Function CmtN_Initialize
	Set CmtN_Config=New TConfig
	CmtN_Config.Load "CmtN_Config"
	If CmtN_Config.Exists("ver")=False Then
		CmtN_Config.Write "CmtN_Charset", "GB2312"
		CmtN_Config.Write "CmtN_MailToAddress", "haphic@126.com"
		CmtN_Config.Write "CmtN_MailReplyToAddress", "haphic@126.com"
		CmtN_Config.Write "CmtN_MailFromName", "博客留言提醒"
		CmtN_Config.Write "CmtN_NotifyCmtLeaver", True
		CmtN_Config.Write "CmtN_MailServerName", "smtp.163.com"
		CmtN_Config.Write "CmtN_MailServerUserName", "haphic@163.com"
		CmtN_Config.Write "CmtN_MailServerUserPwd", "TypeYourPasswordHere"
		CmtN_Config.Write "CmtN_MailFromAddress", "haphic@163.com"
		CmtN_Config.Write "CmtN_MailServerAlternate", "haphic:pwd@smtp.qq.com(haphic@qq.com)"
		CmtN_Config.Write "CmtN_MailSendDelay", False
		CmtN_Config.Write "CmtN_MailSendDelayTime", 120
		
		CmtN_Config.Write "ver", "1.0"
		CmtN_Config.Save
	End If
	CmtN_Charset=CmtN_Config.Read("CmtN_Charset")
	CmtN_MailToAddress=CmtN_Config.Read("CmtN_MailToAddress")
	CmtN_MailReplyToAddress=CmtN_Config.Read("CmtN_MailReplyToAddress")
	CmtN_MailFromName=CmtN_Config.Read("CmtN_MailFromName")
	CmtN_NotifyCmtLeaver=CBool(CmtN_Config.Read("CmtN_NotifyCmtLeaver"))
	CmtN_MailServerName=CmtN_Config.Read("CmtN_MailServerName")
	CmtN_MailServerUserName=CmtN_Config.Read("CmtN_MailServerUserName")
	CmtN_MailServerUserPwd=CmtN_Config.Read("CmtN_MailServerUserPwd")
	CmtN_MailFromAddress=CmtN_Config.Read("CmtN_MailFromAddress")
	CmtN_MailServerAlternate=CmtN_Config.Read("CmtN_MailServerAlternate")
	CmtN_MailSendDelay=CBool(CmtN_Config.Read("CmtN_MailSendDelay"))
	CmtN_MailSendDelayTime=CDbl(CmtN_Config.Read("CmtN_MailSendDelayTime"))
End Function

'*********************************************************
' 目的：    发送评论
'*********************************************************
Function CmtN_SendComment(obj)
	Call CmtN_Initialize
	
	If CmtN_MailToAddress="null" Then Exit Function

	If BlogUser.Level=1 Then Exit Function

	Dim inpID,inpName,inpArticle,inpEmail,inpHomePage,inpIP,inpAgent

	inpID=obj.logID
	inpName=obj.Author
	inpArticle=obj.Content
	inpEmail=obj.email
	inpHomePage=obj.homepage

	inpIP=obj.ip
	inpAgent=obj.agent


	If Len(inpHomePage)>0 Then
		If InStr(inpHomePage,"http://")=0 Then inpHomePage="http://" & inpHomePage
	End If

	inpName=TransferHTML(inpName,"[html-format]")
	inpEmail=TransferHTML(inpEmail,"[html-format]")
	inpHomePage=TransferHTML(inpHomePage,"[html-format]")
	inpArticle=TransferHTML(inpArticle,"[html-format]")

	Dim MA : MA = CmtN_MailToAddress

	Dim MT : MT = ""
	Dim MC : MC = CmtN_Template("cmt")

	Dim User
	Dim objRS
	Dim objArticle

	Set objArticle=New TArticle
	If objArticle.LoadInfoByID(inpID) Then
		Set objRS=objConn.Execute("SELECT TOP 1 [comm_ID],[log_ID] FROM [blog_Comment] WHERE ([log_ID]="& inpID &") ORDER BY [comm_ID] DESC")
		If (Not objRS.bof) And (Not objRS.eof) Then
	
			MT = MT & inpName & " 在您的博客 """& ZC_BLOG_NAME &""" 里评论"
	
			MC = Replace(MC,"<#Cmt_Type#>","评论")
			MC = Replace(MC,"<#Cmt_Article/title#>",objArticle.Title)
			MC = Replace(MC,"<#Cmt_Article/url#>",objArticle.Url)
			MC = Replace(MC,"<#Cmt_Article/PostTime#>",objArticle.PostTime)
			Call GetUsers
			For Each User in Users
				If IsObject(User) Then
					If User.ID=objArticle.AuthorID Then
						MC = Replace(MC,"<#Cmt_Article/author/name#>",User.Name)
					End If
				End If
			Next
			MC = Replace(MC,"<#Cmt_Article/author/name#>","")
			MC = Replace(MC,"<#Cmt_Url#>",objArticle.Url & "#cmt" & objRS("comm_ID"))
		End If
		objRS.close
		Set objRS=Nothing

		Set objArticle=Nothing

	End If


	MC = Replace(MC,"<#MAIL_RECEIVER#>",ZC_BLOG_MASTER)
	MC = Replace(MC,"<#BLOG_LINK#>","<a href="""& ZC_BLOG_HOST &""" title="""& ZC_BLOG_SUB_NAME &""" target=""_blank"">"& ZC_BLOG_NAME &"</a>")

	MC = Replace(MC,"<#Cmt_Author/name#>",inpName)
	MC = Replace(MC,"<#Cmt_Author/url#>",inpHomePage)
	MC = Replace(MC,"<#Cmt_Author/email#>",inpEmail)
	MC = Replace(MC,"<#Cmt_Author/IP#>",inpIP)
	MC = Replace(MC,"<#Cmt_Author/agent#>",inpAgent)

	If Len(inpHomePage)> 0 Then
		MC = Replace(MC,"<#Cmt_Author#>","<a href="""& inpHomePage &""" target=""_blank"">"& inpName &"</a>")
	Else
		MC = Replace(MC,"<#Cmt_Author#>",inpName)
	End If

	MC = Replace(MC,"<#Cmt_PostTime#>",GetTime(Now()))
	MC = Replace(MC,"<#Cmt_Content>",TransferHTML(UBBCode(inpArticle,"[link][link-antispam][font][face]"),"[enter][nofollow]"))

	If CmtN_MailSendDelay Then
		Dim tmpRnd : Randomize : tmpRnd=Int(Rnd*1000)
		Dim tmpName : tmpName = Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&tmpRnd
		Dim tmpValue :tmpValue = MT & vbCrlf & MA & vbCrlf & "null" & vbCrlf & MC
		Call SaveToFile(BlogPath & "zb_users/PLUGIN/CmtN/OutGoingMails/"& tmpName &".html",tmpValue,"utf-8",False)
	Else
		Call CmtN_SendMessageViaJamil(MA,"null",CmtN_MailReplyToAddress,CmtN_MailFromName,MT,MC)
	End If

End Function
'*********************************************************

%>