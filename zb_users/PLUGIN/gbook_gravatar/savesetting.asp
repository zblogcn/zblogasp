<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8及以上的版本
'// 插件制作:    大猪 (http://www.izhu.org)
'// 备    注:    大猪
'// 最后修改：   2012-6-27
'// 最后版本:    1.0.0
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

If CheckPluginState("gbook_gravatar")=False Then Call ShowError(48)

'if not isnumeric(Request.Form("str_DZ_IDS_VALUE")) then
'response.write "<script>alert('文章ID必须填写');< /script>"
'response.End()
'end if

	Call gbook_gravatar_Initialize
'	gbook_gravatar_Config.Write "DZ_IDS_VALUE",Request.Form("str_DZ_IDS_VALUE")
	gbook_gravatar_Config.Write "DZ_AVATAR_VALUE",Request.Form("str_DZ_AVATAR_VALUE")
	gbook_gravatar_Config.Write "DZ_WH_VALUE",Request.Form("str_DZ_WH_VALUE")
	gbook_gravatar_Config.Write "DZ_TITLE_VALUE",Request.Form("str_DZ_TITLE_VALUE")
	gbook_gravatar_Config.Write "DZ_COUNT_VALUE",Request.Form("str_DZ_COUNT_VALUE")
	gbook_gravatar_Config.Save
	Set gbook_gravatar_Config=Nothing

Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>
<script type="text/javascript">window.location="setting.asp"</script>
