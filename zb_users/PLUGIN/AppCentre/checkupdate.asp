<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!-- #include file="function.asp"-->
<%

Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("AppCentre")=False Then Call ShowError(48)


Dim s,id
AppCentre_InitConfig
s=disableupdate_theme

Select Case Request.QueryString("act")
Case "dut"

	id=Request.QueryString("id")

	If InStr(s,id & ":")=0 Then
		s=s & id & ":"
	End If
	app_config.Write "DisableUpdateTheme",s
	app_config.Save

	Response.Redirect BlogHost & "zb_system/cmd.asp?act=ThemeMng"
Case "eut"

	id=Request.QueryString("id")

	If InStr(s,id & ":")>0 Then
		s=Replace(s,id & ":","")
	End If
	app_config.Write "DisableUpdateTheme",s
	app_config.Save

	Response.Redirect BlogHost & "zb_system/cmd.asp?act=ThemeMng"
End Select
%>
