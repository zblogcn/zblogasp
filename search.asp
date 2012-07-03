<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    search.asp
'// 开始时间:    2005.02.17
'// 最后修改:    
'// 备    注:    站内搜索
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_function_md5.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

'plugin node
For Each sAction_Plugin_Searching_Begin in Action_Plugin_Searching_Begin
	If Not IsEmpty(sAction_Plugin_Searching_Begin) Then Call Execute(sAction_Plugin_Searching_Begin)
Next


'检查权限
If Not CheckRights("Search") Then Call ShowError(6)

Dim strQuestion
strQuestion=TransferHTML(Request.QueryString("q"),"[nohtml]")

Dim ArtList
Set ArtList=New TArticleList

ArtList.LoadCache

ArtList.template="SEARCH"

If ArtList.Search(strQuestion) Then

	ArtList.Title=ZC_MSG085 + ":" + TransferHTML(strQuestion,"[html-format]")

	ArtList.Build

	Response.Write ArtList.html

End If

'plugin node
For Each sAction_Plugin_Searching_End in Action_Plugin_Searching_End
	If Not IsEmpty(sAction_Plugin_Searching_End) Then Call Execute(sAction_Plugin_Searching_End)
Next

Call System_Terminate()

%><!-- <%=RunTime()%>ms --><%
If Err.Number<>0 then
	Call ShowError(0)
End If
%>