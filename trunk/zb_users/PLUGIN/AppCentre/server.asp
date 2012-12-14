<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
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
%>
<%
Stop
Dim objXmlHttp,strURL,bolPost,str
bolPost=IIf(Request.ServerVariables("REQUEST_METHOD")="POST",True,False)

Set objXmlHttp=Server.CreateObject("MSXML2.ServerXMLHTTP")

Select Case Request.QueryString("act")
	Case "view"
		strURL="view.asp?"
	Case "catalog"
		strURL="catalog.asp?"
	Case Else
		strURL=""
End Select

strURL=strURL & Request.QueryString

strURL=APPCENTRE_URL & strURL
If bolPost Then objXmlhttp.SetRequestHeader "Content-Type","application/x-www-form-urlencoded"
objXmlHttp.Open Request.ServerVariables("REQUEST_METHOD"),strURL
objXmlHttp.Send Request.Form

Dim strResponse
strResponse=objXmlhttp.ResponseText
strResponse=Replace(strResponse,"catalog.asp?","server.asp?act=catalog&")
strResponse=Replace(strResponse,APPCENTRE_URL&"view.asp?","server.asp?act=view&")
strResponse=Replace(strResponse,APPCENTRE_URL&"""","server.asp""")

Response.Write strResponse
%>