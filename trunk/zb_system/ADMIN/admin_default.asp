<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_function_md5.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_event.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->
<%

Call OpenConnect()
Set BlogUser =New TUser
BlogUser.Verify()

Call CheckReference("")

Dim strAct
strAct="admin"

'检查权限
If Not CheckRights(strAct) Then Call ShowError(6)

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<link rel="stylesheet" rev="stylesheet" href="../CSS/admin.css" type="text/css" media="screen" />
	<title>Z-Blog <%=ZC_MSG248%></title>
</head>
<frameset name="all" rows="70,*" framespacing="0" border="0" frameborder="0">
	<frame name="banner" scrolling="no" src="admin_top.asp" noresize="noresize"  marginwidth="0" marginheight="0" frameborder="0">
	<frameset name="content" cols="150,*" framespacing="0"  border="0" frameborder="0">
		<frame name="list" target="main" src="admin_left.asp"  noresize="noresize"  marginwidth="0" marginheight="0" frameborder="0" scrolling="none" id="list">
		<frame name="main" src="../cmd.asp?<%If Request.ServerVariables("QUERY_STRING")="" Then%>act=SiteInfo<%Else Response.Write Request.ServerVariables("QUERY_STRING") End If%>" noresize="noresize" marginwidth="0" marginheight="0" frameborder="0" scrolling="yes">
	</frameset>
	<noframes>
	<body>

	<p></p>

	</body>
	</noframes>
</frameset>
</html>
<%
Call CloseConnect()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>