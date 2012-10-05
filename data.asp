<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    default.asp
'// 开始时间:    2004.07.25
'// 最后修改:    
'// 备    注:    主页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<%
Dim objRs
'On Error Resume NEXT
Dim PS,NowPage
Dim log_id
log_id=199
PS=100
NowPage=2201
Call OpenConnect
Dim AllComs
Test0
Response.Flush()
Test1
Response.Flush()
Test2

Sub Test0
	StarTime=Timer
	AllComs=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Comment] WHERE [log_ID] ="&log_id&" AND [comm_isCheck]=0")(0)
	Response.Write "LogID: " &log_id&"<br/>"
	Response.write "CommentCount: "&AllComs&"<br/>"
	Response.write "PageSize: "&PS&"<br/>"
	Response.write "NowPage: "&NowPage&"<br/>"
	Response.Write "RunTime: " & RunTime&"<br/><br/>"
End Sub

Sub Test1
	'测试代码1
	StarTime=Timer
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source="SELECT * FROM [blog_Comment] WHERE ([log_ID]="&log_id&" AND [comm_isCheck]=0 AND [comm_ParentID]=0)  ORDER BY [comm_PostTime] DESC"
	objRS.Open()
	objRS.PageSize=PS
	objRS.AbsolutePage =NowPage
	Response.Write "---------TEST1---------<br/>"
	Response.Write "Type: PageSize+AbsolutePage<br/>"
	Response.Write "RunTime: "&RunTime &"<br/>"
	Response.Write "Comm_id: "& objRs("comm_id")&"<br/>"
	Response.Write "---------TEST1---------<br/><br/>"
	Set objRs=Nothing
End Sub

Sub Test2
	'测试代码2
	Dim PageSize2
	PageSize2=PS
	StarTime=Timer
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	If PS*NowPage>AllComs Then
		PageSize2=CLng(AllComs Mod PS)
		If PS*(NowPage-1)+PageSize2>AllComs Then Response.Write "Error!":Response.End
	End If
	objRS.Source="SELECT * FROM (SELECT TOP "&PageSize2&" *  FROM (SELECT TOP "&(PS*NowPage)&" * FROM [blog_Comment]  WHERE ([log_ID]="&log_id&" AND [comm_isCheck]=0 AND [comm_ParentID]=0) ORDER BY [comm_id] DESC) As [Test] ORDER BY [comm_id] asc ) As [test] order by [comm_posttime] desc"
	objRS.Open()
	Response.Write "---------TEST2---------<br/>"
	Response.Write "Type: Double Top<br/>"
	Response.Write "PageSize2: "&PageSize2 & "<br/>"
	Response.Write "RunTime: "&RunTime &"<br/>"
	Response.Write "Comm_id: "& objRs("comm_id")&"<br/>"
	Response.Write "---------TEST2---------<br/><br/>"
End Sub
%>