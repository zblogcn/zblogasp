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
If ZC_MSSQL_ENABLE=False Then
	objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_Template] VARCHAR(50) default """"")
	objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_FullUrl] VARCHAR(255) default """"")
Else
	objConn.execute("ALTER TABLE [blog_Member] ADD [mem_Template] nvarchar(50) default ''")
	objConn.execute("ALTER TABLE [blog_Member] ADD [mem_FullUrl] nvarchar(255) default ''")
End If
%>
