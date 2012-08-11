<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8及以上的版本
'// 插件制作:    大猪 (http://www.izhu.org)
'// 备    注:    大猪滔滔
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

If CheckPluginState("dztaotao")=False Then Call ShowError(48)

	Call dztaotao_Initialize
	dztaotao_Config.Write "DZTAOTAO_TITLE_VALUE",Request.Form("strZC_DZTAOTAO_TITLE_VALUE")
	dztaotao_Config.Write "DZTAOTAO_RELEASE_VALUE",Request.Form("strZC_DZTAOTAO_RELEASE_VALUE")
	dztaotao_Config.Write "DZTAOTAO_PAGECOUNT_VALUE",Request.Form("strDZTAOTAO_PAGECOUNT_VALUE")
	dztaotao_Config.Write "DZTAOTAO_PAGEWIDTH_VALUE",Request.Form("strDZTAOTAO_PAGEWIDTH_VALUE")
	dztaotao_Config.Write "DZTAOTAO_CHK_VALUE",Request.Form("strDZTAOTAO_CHK_VALUE")
	dztaotao_Config.Write "DZTAOTAO_CMTCHK_VALUE",Request.Form("strDZTAOTAO_CMTCHK_VALUE")	
	dztaotao_Config.Write "DZTAOTAO_CMTLIMIT_VALUE",Request.Form("strDZTAOTAO_CMTLIMIT_VALUE")	
	dztaotao_Config.Write "DZTAOTAO_ISIMG_VALUE",Request.Form("strZC_DZTAOTAO_ISIMG_VALUE")
	dztaotao_Config.Save
	Set dztaotao_Config=Nothing

Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>
<script type="text/javascript">window.location="setting.asp"</script>
