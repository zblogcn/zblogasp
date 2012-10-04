<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    view.asp
'// 开始时间:    2004.07.30
'// 最后修改:    
'// 备    注:    查看页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

'plugin node
For Each sAction_Plugin_View_Begin in Action_Plugin_View_Begin
	If Not IsEmpty(sAction_Plugin_View_Begin) Then Call Execute(sAction_Plugin_View_Begin)
Next

Dim objRS
Dim Article
Set Article=New TArticle
Dim Config
Set Config=New TConfig
Config.Load "GuestBook"
Dim j
j=Config.Read("g")
If j<>"" Then
	j=CLng(j)
	If j=0 Then Response.End 
Else
	Response.End 
End If

If Article.LoadInfoByID(j) Then

	If Article.Export(ZC_DISPLAY_MODE_ALL)= True Then
		Article.Build
		Response.Write Article.html
	End If

End If

'plugin node
For Each sAction_Plugin_View_End in Action_Plugin_View_End
	If Not IsEmpty(sAction_Plugin_View_End) Then Call Execute(sAction_Plugin_View_End)
Next

Call System_Terminate()

%>
<!-- <%=RunTime()%>ms --><%
If Err.Number<>0 then
	Call ShowError(0)
End If
%>