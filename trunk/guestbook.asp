<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd) & Sipo
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    guestbook.asp
'// 开始时间:    2007-01-03
'// 最后修改:    
'// 备    注:    GuestBook
'///////////////////////////////////////////////////////////////////////////////
%>
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

'plugin node
For Each sAction_Plugin_Guestbook_Begin in Action_Plugin_Guestbook_Begin
	If Not IsEmpty(sAction_Plugin_Guestbook_Begin) Then Call Execute(sAction_Plugin_Guestbook_Begin)
Next

Dim Page
Set Page=New TArticle

If Page.LoadInfoByID(ZC_GUESTBOOK_ID) Then

	If Page.Level=1 Then Call ShowError(9)
	If Page.Level=2 Then
		If Not CheckRights("Root") Then
			If (Article.AuthorID<>BlogUser.ID) Then Call ShowError(6)
		End If
	End If

	Page.template="SINGLE"
	If Page.Export(ZC_DISPLAY_MODE_ALL)= True Then
		Page.Build
		Response.Write Page.html
	End If

End If


'plugin node
For Each sAction_Plugin_Guestbook_End in Action_Plugin_Guestbook_End
	If Not IsEmpty(sAction_Plugin_Guestbook_End) Then Call Execute(sAction_Plugin_Guestbook_End)
Next

Call System_Terminate()

%><!-- <%=RunTime()%>ms --><%
If Err.Number<>0 then
	Call ShowError(0)
End If
%>