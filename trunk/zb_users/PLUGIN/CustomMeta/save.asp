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
If CheckPluginState("CustomMeta")=False Then Call ShowError(48)

Dim c
Set c=New TConfig
c.Load "CustomMeta"


Dim m,i
Dim name,note
Set m=New TMeta

name=Split(Request.Form("MetaName"),", ")
note=Split(Request.Form("MetaNote"),", ")

For i=LBound(name) To UBound(name)
m.SetValue name(i),note(i)
Next

If Right(Request.ServerVariables("HTTP_REFERER"),8)="main.asp" Then

	c.Write "LogMeta",m.SaveString
	c.Save

	SetBlogHint_Custom("配置已保存!")
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End If

If Right(Request.ServerVariables("HTTP_REFERER"),8)="cate.asp" Then

	c.Write "CateMeta",m.SaveString
	c.Save

	SetBlogHint_Custom("配置已保存!")
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End If

If Right(Request.ServerVariables("HTTP_REFERER"),8)="user.asp" Then

	c.Write "UserMeta",m.SaveString
	c.Save

	SetBlogHint_Custom("配置已保存!")
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End If

%>