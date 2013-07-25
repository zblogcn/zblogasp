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
	
	
	
If Request.QueryString("act")="delall" Then
	Totoro_Config.Delete
	Call SetBlogHint_Custom("配置初始化成功")
Else
	
	Totoro_Config.Write "TOTORO_INTERVAL_VALUE",Request.Form("strZC_TOTORO_INTERVAL_VALUE")
	Totoro_Config.Write "TOTORO_BADWORD_VALUE",Request.Form("strZC_TOTORO_BADWORD_VALUE")
	Totoro_Config.Write "TOTORO_HYPERLINK_VALUE",Request.Form("strZC_TOTORO_HYPERLINK_VALUE")
	Totoro_Config.Write "TOTORO_NAME_VALUE",Request.Form("strZC_TOTORO_NAME_VALUE")
	Totoro_Config.Write "TOTORO_LEVEL_VALUE",Request.Form("strZC_TOTORO_LEVEL_VALUE")	
	Totoro_Config.Write "TOTORO_SV_THRESHOLD",Request.Form("strZC_TOTORO_SV_THRESHOLD")	
	Totoro_Config.Write "TOTORO_DEL_DIRECTLY",Request.Form("bolTOTORO_DEL_DIRECTLY")
	Totoro_Config.Write "TOTORO_ConHuoxingwen",Request.Form("bolTOTORO_ConHuoxingwen")
	
	Totoro_Config.Write "TOTORO_SV_THRESHOLD2",Request.Form("strZC_TOTORO_SV_THRESHOLD2")
	Totoro_Config.Write "TOTORO_NUMBER_VALUE",Request.Form("strTOTORO_NUMBER_VALUE")
	
	Totoro_Config.Write "TOTORO_REPLACE_KEYWORD",Request.Form("strZC_TOTORO_REPLACE_KEYWORD")
	Totoro_Config.Write "TOTORO_CHINESESV",Request.Form("strZC_TOTORO_CHINESESV")
	Totoro_Config.Write "TOTORO_KILLIP",Request.Form("strZC_TOTORO_KILLIP")
	Totoro_Config.Write "TOTORO_FILTERIP",Request.Form("strTOTORO_FILTERIP")
	Totoro_Config.Write "TOTORO_PM",Request.Form("bolTOTORO_PM")
	Totoro_Config.Write "TOTORO_TRANTOSIMP",Request.Form("bolTOTORO_TRANTOSIMP")
	Totoro_Config.Write "TOTORO_THROWSTR",Replace(Replace(Request.Form("strTOTORO_THROWSTR"),vbCrlf,""),vbLf,"")
	Totoro_Config.Write "TOTORO_KILLIPSTR",Replace(Replace(Request.Form("strTOTORO_KILLIPSTR"),vbCrlf,""),vbLf,"")
	Totoro_Config.Write "TOTORO_CHECKSTR",Replace(Replace(Request.Form("strTOTORO_CHECKSTR"),vbCrlf,""),vbLf,"")
	
	Dim strData
	strData=Replace(Replace(Request.Form("strZC_TOTORO_BADWORD_LIST"),vbCrlf,""),vbLf,"")
	If strData<>"" Then
		If CheckRegExp(Left(strData,1),"[a-zA-Z0-9]") Then strData=vbsescape("插件防BUG黑词|") & strData
	End If
	Totoro_Config.Write "TOTORO_BADWORD_LIST",strData
	
	strData=Replace(Replace(Request.Form("strZC_TOTORO_REPLACE_LIST"),vbCrlf,""),vbLf,"")
	If strData<>"" Then
		If CheckRegExp(Left(strData,1),"[a-zA-Z0-9]") Then strData=vbsescape("插件防BUG黑词|") & strData
	End If
	Totoro_Config.Write "TOTORO_REPLACE_LIST",strData
	
	
	Totoro_Config.Save
	Set Totoro_Config=Nothing
	Call SetBlogHint_Custom("配置保存成功！强烈建议去<a href='onlinetest.asp'>对配置进行一次测试！</a>")
	
End If
Response.Redirect "setting.asp"
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>