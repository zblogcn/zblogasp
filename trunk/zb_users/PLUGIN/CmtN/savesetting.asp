<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    haphic
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 
Call CmtN_Initialize
	Response.Write "<div style=""position:absolute;top:1px;right:1px;height:15px;width:180px;margin:0;padding:5px 10px;background:#8B0000;color:#FFFFFF;font-size:12px;"">Saving Data, Please Wait ...</div>"
	Response.Flush

	Dim strContent
	strContent=LoadFromFile(BlogPath & "/PLUGIN/CmtN/config.asp","utf-8")

	Dim strCmtN_MailServerName
	strCmtN_MailServerName=Request.Form("strCmtN_MailServerName")
	If Not IsEmpty(strCmtN_MailServerName) Then
		CmtN_Config.Write "CmtN_MailServerName",IIf(strCmtN_MailServerName=False,"False",strCmtN_MailServerName)
	End If

	Dim strCmtN_MailServerAlternate
	strCmtN_MailServerAlternate=Request.Form("strCmtN_MailServerAlternate")
	If Not IsEmpty(strCmtN_MailServerAlternate) Then
		CmtN_Config.Write "CmtN_MailServerAlternate",IIf(strCmtN_MailServerAlternate=False,"False",strCmtN_MailServerAlternate)
	End If

	Dim strCmtN_MailServerUserName
	strCmtN_MailServerUserName=Request.Form("strCmtN_MailServerUserName")
	If Not IsEmpty(strCmtN_MailServerUserName) Then
		CmtN_Config.Write "CmtN_MailServerUserName",IIf(strCmtN_MailServerUserName=False,"False",strCmtN_MailServerUserName)
	End If

	Dim strCmtN_MailServerUserPwd
	strCmtN_MailServerUserPwd=Request.Form("strCmtN_MailServerUserPwd")
	If Not IsEmpty(strCmtN_MailServerUserPwd) Then
		CmtN_Config.Write "CmtN_MailServerUserPwd",IIf(strCmtN_MailServerUserPwd=False,"False",strCmtN_MailServerUserPwd)
	End If

	Dim strCmtN_MailFromName
	strCmtN_MailFromName=Request.Form("strCmtN_MailFromName")
	If Not IsEmpty(strCmtN_MailFromName) Then
		CmtN_Config.Write "CmtN_MailFromName",IIf(strCmtN_MailFromName=False,"False",strCmtN_MailFromName)
	End If

	Dim strCmtN_MailFromAddress
	strCmtN_MailFromAddress=Request.Form("strCmtN_MailFromAddress")
	If Not IsEmpty(strCmtN_MailFromAddress) Then
		CmtN_Config.Write "CmtN_MailFromAddress",IIf(strCmtN_MailFromAddress=False,"False",strCmtN_MailFromAddress)
	End If

	Dim strCmtN_MailReplyToAddress
	strCmtN_MailReplyToAddress=Request.Form("strCmtN_MailReplyToAddress")
	If Not IsEmpty(strCmtN_MailReplyToAddress) Then
		CmtN_Config.Write "CmtN_MailReplyToAddress",IIf(strCmtN_MailReplyToAddress=False,"False",strCmtN_MailReplyToAddress)
	End If

	Dim strCmtN_MailToAddress
	strCmtN_MailToAddress=Request.Form("strCmtN_MailToAddress")
	If Not IsEmpty(strCmtN_MailToAddress) Then
		If Not InStr(strCmtN_MailToAddress,"@")>0 Then strCmtN_MailToAddress="null"
		CmtN_Config.Write "CmtN_MailToAddress",IIf(strCmtN_MailToAddress=False,"False",strCmtN_MailToAddress)
	End If

	Dim strCmtN_MailSendDelayTime
	strCmtN_MailSendDelayTime=Request.Form("strCmtN_MailSendDelayTime")
	If Not IsEmpty(strCmtN_MailSendDelayTime) Then
		CmtN_Config.Write "CmtN_MailSendDelayTime",IIf(strCmtN_MailSendDelayTime=False,"False",strCmtN_MailSendDelayTime)
	End If

	Dim strCmtN_MailSendDelay
	strCmtN_MailSendDelay=Request.Form("strCmtN_MailSendDelay")
	CmtN_Config.Write "CmtN_MailSendDelay",IIf(strCmtN_MailSendDelay="True","True","False")
	
	Dim strCmtN_NotifyCmtLeaver
	strCmtN_NotifyCmtLeaver=Request.Form("strCmtN_NotifyCmtLeaver")
	CmtN_Config.Write "CmtN_NotifyCmtLeaver",IIf(strCmtN_NotifyCmtLeaver="True","True","False")


	CmtN_Config.Save


Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>
<%If Request.Form("TestMail")="True" Then%>
<script type="text/javascript">window.location="setting.asp?act=TestMail"</script>
<%Else%>
<script type="text/javascript">window.location="setting.asp"</script>
<%End If%>
