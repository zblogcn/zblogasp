<%@ CODEPAGE=65001 %>
<%
' ***********************************************************************************
'    如果您看到了这个提示，那么我们很遗憾地通知您，您的空间不支持 ASP 。
'    If you see this notice, we regret to inform you that 
'    your web hosting service doesn't support ASP so Z-Blog can't run on it.

'    也就是说，您的空间可能是静态空间或 PHP 空间，或未在 IIS 内安装 ASP 组件。
'    It means that you may have a web hosting service supporting only static resources or installed PHP.
'    If you're using IIS, maybe you don’t have ASP Extension installed.

'    推荐您：
'    Recommend you:

'            > 下载并安装Z-BlogPHP  > http://www.zblogcn.com/zblogphp/
'            > Try Z-BlogPHP > http://www.zblogcn.com/zblogphp/

'    如果您仍然需要使用Z-BlogASP：
'    Still need Z-BlogASP?

'            > 联系空间商，更换空间为支持ASP的空间
'            > Contact your provider, and let them provice a new hosting which supports ASP.

'            > 打开 IIS 的 ASP 组件
'            > Install ASP Extension on IIS
'
' ***********************************************************************************
 
%>
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
<% On Error Resume Next %>
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
Dim html

'plugin node
For Each sAction_Plugin_Default_WithOutConnect_Begin in Action_Plugin_Default_WithOutConnect_Begin
	If Not IsEmpty(sAction_Plugin_Default_WithOutConnect_Begin) Then Call Execute(sAction_Plugin_Default_WithOutConnect_Begin)
Next

If ZC_DATABASE_PATH="" And ZC_MSSQL_DATABASE="" Then Response.Redirect("zb_install/")

If ZC_HTTP_LASTMODIFIED=True Then
	Response.AddHeader "Last-Modified",ParseDateForRFC822GMT(GetFileModified(BlogPath & "zb_users\cache\default.asp"))
End If

html=LoadFromFile(BlogPath & "zb_users\cache\default.asp","utf-8")

If Len(html)>0 Then
	Response.Write Replace(html,"<#ZC_BLOG_HOST#>",BlogHost)
	Response.Write "<!-- " & RunTime() & "ms -->"
	Response.End
End If

Call System_Initialize()

'plugin node
For Each sAction_Plugin_Default_Begin in Action_Plugin_Default_Begin
	If Not IsEmpty(sAction_Plugin_Default_Begin) Then Call Execute(sAction_Plugin_Default_Begin)
Next


Dim ArtList
Set ArtList=New TArticleList

If ArtList.Export("","","","","",ZC_DISPLAY_MODE_INTRO) Then

	ArtList.Build

	html=ArtList.html

	Response.Write html

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