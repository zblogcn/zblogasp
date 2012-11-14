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
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%
Call ActivePlugin

'plugin node
For Each sAction_Plugin_Default_Begin in Action_Plugin_Default_Begin
	If Not IsEmpty(sAction_Plugin_Default_Begin) Then Call Execute(sAction_Plugin_Default_Begin)
Next


If CheckMobile() Then Response.Redirect ZC_FILENAME_WAP

If ZC_DATABASE_PATH="" And ZC_MSSQL_DATABASE="" Then Response.Redirect("zb_install/")

If ZC_HTTP_LASTMODIFIED=True Then
	Response.AddHeader "Last-Modified",GetFileModified(BlogPath & "zb_users\cache\default.html")
End If

Dim s
s=LoadFromFile(BlogPath & "zb_users\cache\default.html","utf-8")

If Len(s)>0 Then
	Response.Write Replace(s,"<#ZC_BLOG_HOST#>",BlogHost)
	Response.Write "<!-- " & RunTime() & "ms -->"
	Response.End
End If


Call System_Initialize()


Dim ArtList
Set ArtList=New TArticleList

If ArtList.Export("","","","","",ZC_DISPLAY_MODE_INTRO) Then

	ArtList.Build

	Response.Write ArtList.html

End If

'plugin node
For Each sAction_Plugin_Default_End in Action_Plugin_Default_End
	If Not IsEmpty(sAction_Plugin_Default_End) Then Call Execute(sAction_Plugin_Default_End)
Next

Call System_Terminate()


If Err.Number<>0 then
	Call ShowError(0)
End If
%>