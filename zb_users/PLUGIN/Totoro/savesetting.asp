<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.5及以上的版本
'// 插件制作:    williamlong(http://www.williamlong.info)
'// 备    注:    反垃圾留言的插件代码
'// 最后修改：   2006-6-27
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

If CheckPluginState("Totoro")=False Then Call ShowError(48)

	Call Totoro_Initialize
	Totoro_Config.Write "TOTORO_INTERVAL_VALUE",Request.Form("strZC_TOTORO_INTERVAL_VALUE")
	Totoro_Config.Write "TOTORO_BADWORD_VALUE",Request.Form("strZC_TOTORO_BADWORD_VALUE")
	Totoro_Config.Write "TOTORO_HYPERLINK_VALUE",Request.Form("strZC_TOTORO_HYPERLINK_VALUE")
	Totoro_Config.Write "TOTORO_NAME_VALUE",Request.Form("strZC_TOTORO_NAME_VALUE")
	Totoro_Config.Write "TOTORO_LEVEL_VALUE",Request.Form("strZC_TOTORO_LEVEL_VALUE")	
	Totoro_Config.Write "TOTORO_SV_THRESHOLD",Request.Form("strZC_TOTORO_SV_THRESHOLD")	
	Totoro_Config.Write "TOTORO_DEL_DIRECTLY",IIf(IsEmpty(Request.Form("bolTOTORO_DEL_DIRECTLY")),"False","True")
	Totoro_Config.Write "TOTORO_ConHuoxingwen",IIf(IsEmpty(Request.Form("bolTOTORO_ConHuoxingwen")),"False","True")
	Totoro_Config.Write "TOTORO_BADWORD_LIST",Replace(Replace(Request.Form("strZC_TOTORO_BADWORD_LIST"),vbCrlf,""),vbLf,"")
	Totoro_Config.Write "TOTORO_SV_THRESHOLD2",Request.Form("strZC_TOTORO_SV_THRESHOLD2")
	Totoro_Config.Write "TOTORO_NUMBER_VALUE",Request.Form("strTOTORO_NUMBER_VALUE")
	Totoro_Config.Write "TOTORO_REPLACE_LIST",Replace(Replace(Request.Form("strZC_TOTORO_REPLACE_LIST"),vbCrlf,""),vbLf,"")
	Totoro_Config.Write "TOTORO_REPLACE_KEYWORD",Request.Form("strZC_TOTORO_REPLACE_KEYWORD")
	Totoro_Config.Write "TOTORO_CHINESESV",Request.Form("strZC_TOTORO_CHINESESV")
	Totoro_Config.Save
	Set Totoro_Config=Nothing

Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>
<script type="text/javascript">window.location="setting.asp"</script>
