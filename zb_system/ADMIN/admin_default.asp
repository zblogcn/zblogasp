<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
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
BlogTitle="Z-Blog后台管理"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
<!--#include file="admin_header.asp"-->
</head>
<body>
<!--#include file="admin_top.asp"-->
<div id="main">
<div class="main_right">
  <div class="yui">
    <div class="content">
      <div class="wrapper"> 主要内容区，这个区域内的板块背景色为白色。比如： </div>
    </div>
  </div></div>
<!--#include file="admin_left.asp"-->

</div>
</body>
</html>
<%
Call CloseConnect()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
