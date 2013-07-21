<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("STACentre")=False Then Call ShowError(48)

BlogTitle="静态中心配置插件"


Dim a,b,c,d

Set d=CreateObject("Scripting.Dictionary")

For Each a In Request.Form 
	b=Mid(a,4,Len(a))
	If BlogConfig.Exists(b)=True Then
		d.add b,Request.Form(a)
	End If
Next

If d.Exists("ZC_ARTICLE_REGEX") Then If BlogConfig.Read("ZC_ARTICLE_REGEX")<>d.Item("ZC_ARTICLE_REGEX")Then Call SetBlogHint(Empty,Empty,True)
If d.Exists("ZC_PAGE_REGEX") Then If BlogConfig.Read("ZC_PAGE_REGEX")<>d.Item("ZC_PAGE_REGEX")Then Call SetBlogHint(Empty,Empty,True)
If d.Exists("ZC_STATIC_MODE") Then If BlogConfig.Read("ZC_STATIC_MODE")<>d.Item("ZC_STATIC_MODE")Then Call SetBlogHint(Empty,Empty,True)
If d.Exists("ZC_POST_STATIC_MODE") Then If BlogConfig.Read("ZC_POST_STATIC_MODE")<>d.Item("ZC_POST_STATIC_MODE")Then Call SetBlogHint(Empty,Empty,True)

For Each a In d.Keys
	Call BlogConfig.Write(a,d.Item(a))
Next

Call SaveConfig2Option()
Call SetBlogHint(True,True,Empty)
If d.Item("ZC_STATIC_MODE")="REWRITE" Then
	SetBlogHint_Custom("设置已保存！请选择你的伪静态组件以生成规则！")
	Response.Redirect "list.asp"'Request.ServerVariables("HTTP_REFERER")
Else
	Response.Redirect "main.asp"
End If
%>