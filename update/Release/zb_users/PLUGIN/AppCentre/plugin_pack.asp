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
<!-- #include file="function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("AppCentre")=False Then Call ShowError(48)



Dim ZipPathDir,ZipPathFile,Pack_PluginDir,ID

ID=Request.QueryString("id")

ZipPathDir = BlogPath & "zb_users\plugin\" & ID & "\"
ZipPathFile = BlogPath & "zb_users\cache\" & MD5(ZC_BLOG_CLSID & ID) & ".zba"
Pack_PluginDir = ID & "\"


Call CreatePluginXml(ZipPathFile)
Call LoadAppFiles(ZipPathDir,ZipPathFile,Pack_PluginDir)


Response.AddHeader   "Content-Disposition","attachment;filename="& ID &".zba"
Response.ContentType = "application/octet-stream"

Server.Transfer "../../cache/"& MD5(ZC_BLOG_CLSID & ID) &".zba"


%>