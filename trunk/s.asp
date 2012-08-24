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
If Not CheckUpdateDB("[coun_Content]","[blog_Counter]") Then
	IF ZC_MSSQL_ENABLE=True Then	
			objConn.execute("ALTER TABLE [blog_Counter] ADD coun_Content ntext default ''")
			objConn.execute("ALTER TABLE [blog_Counter] ADD coun_UserID int default 0")
			objConn.execute("ALTER TABLE [blog_Counter] ADD coun_PostData ntext default ''")
			objConn.execute("ALTER TABLE [blog_Counter] ADD coun_URL ntext default ''")
			objConn.execute("ALTER TABLE [blog_Counter] ADD coun_AllRequestHeader ntext default '' ")
	ELSE
			objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_Content text default """"")
			objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_UserID int default 0")
			objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_PostData  text default """"")
			objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_URL  text default """"")
			objConn.execute("ALTER TABLE [blog_Counter] ADD COLUMN coun_AllRequestHeader  text default """"")
	
	End IF
End If
Dim a
Set a=New TCounter
a.Add "test",False

echo a.ID
echo a.ip
echo a.referer
echo a.agent
echo a.AllRequestHeader
echo a.postdata
echo a.userid
echo a.posttime
echo a.url
echo a.content
Function echo(s)
	Response.write s&"<br/>"
End FUnction
%>