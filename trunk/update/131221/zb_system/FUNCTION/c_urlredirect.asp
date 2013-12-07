<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)&(sipo)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_urlredirect.asp
'// 开始时间:    2007-1-24
'// 最后修改:    
'// 备    注:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="c_function.asp" -->
<%

Dim strReferer
strReferer=CStr(Request.ServerVariables("HTTP_REFERER"))

If Instr(strReferer,GetCurrentHost())=0 Then 
	ShowError(5)
End If


Dim strUrl
strUrl=URLDecodeForAntiSpam(Request.QueryString("url"))
If strUrl="" Then strUrl=GetCurrentHost()

Response.Redirect strUrl

%>