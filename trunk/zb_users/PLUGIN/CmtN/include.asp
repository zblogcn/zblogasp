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


Dim CmtN


'注册插件
Call RegisterPlugin("CmtN","ActivePlugin_CmtN")


'具体的接口挂接
Function ActivePlugin_CmtN() 

	'挂上接口
	'Action_Plugin_CommentPost_Succeed
	Call Add_Filter_Plugin("Filter_Plugin_PostComment_Succeed","CmtN_SendComment")

	'挂上接口
	'Action_Plugin_Catalog_End
	Call Add_Action_Plugin("Action_Plugin_Catalog_End","CmtN_SendOutGoingMails")
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
	Set CmtN=New CmtN_Class
End Function

'*********************************************************
' 目的：    发送评论
'*********************************************************
Function CmtN_SendComment(obj)
	Call CmtN_Initialize
	'If CmtN_MailToAddress="null" Then Exit Function

	'If BlogUser.Level=1 Then Exit Function
	If obj.ParentID=0 Then CmtN.Template="cmt" Else Cmtn.Template="rev"
	CmtN.MakeCommentTemplate obj,False
	If obj.ParentID=0 Then 
		CmtN.mailTo=obj.email
	Else
		CmtN.mailTo=CmtN_MailToAddress
	End if
	If CmtN_MailSendDelay Then
		Dim tmpRnd : Randomize : tmpRnd=Int(Rnd*1000)
		Dim tmpName : tmpName = Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&tmpRnd
		Dim tmpValue :tmpValue = cmtn.mailSubject & vbCrlf & CmtN.mailTo & vbCrlf & "null" & vbCrlf & CmtN.mailBody
		Call SaveToFile(BlogPath & "zb_users/PLUGIN/CmtN/OutGoingMails/"& tmpName &".html",tmpValue,"utf-8",False)
	Else
		CmtN.mailReply=CmtN_MailReplyToAddress
		CmtN.mailName=CmtN_MailFromName
		CmtN.mailTo2=CmtN_MailToAddress
		CmtN.Send
	End If

End Function
'*********************************************************



%>