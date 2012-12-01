<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷(YT.Single)
'// 技术支持:    33195@qq.com
'// 程序名称:    	YT.Build
'// 开始时间:    	2011.03.26
'// 最后修改:    2012.08.24
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% Server.ScriptTimeOut=10000 %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #INCLUDE FILE="../../C_OPTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_FUNCTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_LIB.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_BASE.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_EVENT.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_PLUGIN.ASP" -->
<!-- #INCLUDE FILE="../../PLUGIN/P_CONFIG.ASP" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("YTBuild")=False Then Call ShowError(48)
Dim bl
Set bl = new YTBuildLib
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	Response.ContentType = "application/json"
	If ZC_STATIC_MODE = "REWRITE" Then
		Select Case Request.Form("Act")
			Case "Default"
				Response.Write(LCase(bl.Default()))
			Case "ThreadView"
				Response.Write(LCase(bl.ThreadView(Request.Form("ID"))))
			Case "ThreadCatalog"
				Response.Write(LCase(bl.ThreadCatalog(Request.Form("Key"),Request.Form("ID"),Request.Form("Page"))))
			Case "View"
				Response.Write(LCase(bl.View(Request.Form("Key"),Request.Form("aC"))))
			Case "Catalog"
				Response.Write(bl.Catalog(Request.Form("Key"),Request.Form("aC")))
		End Select
	Else
		Response.Write("未启用静态功能")
	End If
End If
Set bl = Nothing
%>