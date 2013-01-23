<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ZBDK")=False Then Call ShowError(48)
BlogTitle=zbdk_title
Dim o
o=MsgCount

Select Case Request.QueryString("id")
	Case "language"
	For MsgCount=1 To o
		Response.Write "<tr height='32'><td><input type='text' value='&lt;#ZC_MSG"&Right("000"&MsgCount,3)&"#&gt;' style='width:100%'/></td><td>"
		Execute("Response.Write ZC_MSG" & Right("000"&MsgCount,3))
		Response.Write "</td><td></td></tr>"
	Next
End Select
%>
