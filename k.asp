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
<%Function CheckUpdateDB(a,b)
	Err.Clear
	On Error Resume Next
	Dim Rs
	Set Rs=objConn.execute("SELECT "&a&" FROM "&b)
	Set Rs=Nothing
	If Err.Number=0 Then
	CheckUpdateDB=True
	Else
	Err.Clear
	CheckUpdateDB=False
	End If	
End Function

Call System_Initialize()

If Not CheckUpdateDB("[mem_Url]","[blog_Member]") Then
	IF ZC_MSSQL_ENABLE=True Then	
		objConn.execute("ALTER TABLE [blog_Member] ADD [mem_Url] nvarchar(255) default '' ")
	ELSE
		objConn.execute("ALTER TABLE [blog_Member] ADD COLUMN [mem_Url]  VARCHAR(255) default """"")
	End IF
End If

response.write "ok"
%>