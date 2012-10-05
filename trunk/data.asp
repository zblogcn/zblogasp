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
PS=100
NowPage=1
Call OpenConnect


'测试代码1
StarTime=Timer
Set objRS=Server.CreateObject("ADODB.Recordset")
objRS.CursorType = adOpenKeyset
objRS.LockType = adLockReadOnly
objRS.ActiveConnection=objConn
objRS.Source="SELECT * FROM [blog_Comment] WHERE ([log_ID]=199 AND [comm_isCheck]=0 AND [comm_ParentID]=0)  ORDER BY [comm_PostTime] DESC"
objRS.Open()
objRS.PageSize=PS
objRS.AbsolutePage =NowPage
Response.Write RunTime & "&nbsp;" & objRs("comm_id")
Set objRs=Nothing


'测试代码2
StarTime=Timer
Set objRS=Server.CreateObject("ADODB.Recordset")
Dim Page1
Page1=objConn.Execute ("select top 1 [comm_id] from (SELECT TOP "&PS&" [comm_id] FROM [BLOG_COMMENT]  WHERE ([log_ID]=199 AND [comm_isCheck]=0 AND [comm_ParentID]=0) order by [comm_id] asc) as [a] order by [comm_id] desc")(0)
objRS.CursorType = adOpenKeyset
objRS.LockType = adLockReadOnly
objRS.ActiveConnection=objConn
objRS.Source="SELECT * FROM (SELECT TOP "&PS&" *  FROM (SELECT TOP "&(PS*NowPage)&" * FROM [blog_Comment]  WHERE ([log_ID]=199 AND [comm_isCheck]=0 AND [comm_ParentID]=0) ORDER BY [comm_id] DESC) As [Test] ORDER BY [comm_id] asc ) As [test] order by [comm_posttime] desc"
objRS.Open()
If objRs("comm_id")=Page1 And NowPage<>1 Then Response.Write "err"':Response.END
Response.Write "<br/>"&RunTime & "&nbsp;" & objRs("comm_id")
%>