<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../../zb_system/admin/ueditor/asp/aspincludefile.asp"-->
<!-- #include file="function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("AppCentre")=False Then Call ShowError(48)

Dim strRnd
Randomize
strRnd=Rnd
Dim strURL
strURL=Request.QueryString("url")
If Left(strURL,Len(APPCENTRE_URL))=APPCENTRE_URL Then 
	Randomize
	Dim objXmlHttp
	Set objXmlHttp=Server.CreateObject("msxml2.serverxmlhttp")
	objXmlhttp.Open "GET",strURL & "?" & strRnd
	objXmlHttp.Send
	Call SaveBinary(objXmlhttp.ResponseBody,BlogPath&"zb_users\cache\temp_" & strRnd & ".zba")
	Call InstallApp(BlogPath&"zb_users\cache\temp_" & strRnd & ".zba")
	Call DelToFile(BlogPath&"zb_users\cache\temp_" & strRnd & ".zba")
	Response.Redirect BlogHost & "zb_system/cmd.asp?act=PlugInMng"
Else
	Response.Write "Illegal URL!"
End If
%>