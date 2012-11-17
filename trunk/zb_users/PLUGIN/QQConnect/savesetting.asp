<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize
init_qqconnect()
Call CheckReference("")
If CheckPluginState("QQConnect")=False Then Call ShowError(48)
If BlogUser.Level=1 Then 
	Dim b,a
	Set a=qqconnect.tconfig
	For b=97 To 105
		a.Write Chr(b),Request.Form(Chr(b))
	Next
	a.Write "AppID",Request.Form("AppID")
	a.Write "KEY",Request.Form("Key")
	a.Write "a1",Request.Form("a1")
	a.Write "Gravatar",Request.Form("Gravatar")
	a.Write "content",Request.Form("content")
	a.Write "pl",Request.Form("pl")
	a.Save
End If
If BlogUser.Level<5 Then
	BlogUser.Meta.SetValue "qqconnect_sync",Request.Form("synctoqzone")
	BlogUser.Edit BlogUser
End If
Call SetBlogHint(True,Empty,Empty)
Response.Redirect "setting.asp"
%>