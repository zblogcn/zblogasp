<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
If CheckPluginState("HeartVote")=False Then Response.End

GetReallyDirectory()

Dim id


Dim allvote
Dim alluser

allvote=0
alluser=0

id=CInt(Request.QueryString("id"))


Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BlogPath & "zb_users/plugin/heartvote/db.asp"


Dim objRS
Set objRS=Server.CreateObject("ADODB.Recordset")
objRS.CursorType = adOpenKeyset
objRS.LockType = adLockReadOnly
objRS.ActiveConnection=objConn
objRS.Source=""


objRS.Open("SELECT SUM([vote])AS allvote,COUNT([ip]) AS alluser FROM [vote] WHERE [aid]=" & id)

If (Not objRS.bof) And (Not objRS.eof) Then

	alluser=objRS("alluser")
	allvote=objRS("allvote")
	allvote=allvote\alluser

End If

objRS.Close
Set objRS=Nothing

objConn.Close

response.write "showVote('"&allvote&"','"&alluser&"')"

%>