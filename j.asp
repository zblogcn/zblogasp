<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

objConn.execute("ALTER TABLE [blog_Article] ADD [log_Type] int default 0")

objConn.execute("UPDATE [blog_Article] SET [log_Type]=0")

objConn.execute("UPDATE [blog_Article] SET [log_Type]=1 WHERE [log_CateID]=0")


%>
<br/><%=RunTime()%>ms<br/>
<%

Call System_Terminate()

%>