<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
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

Dim ZipPathFile
ZipPathFile=BlogPath & "zb_users\cache\temp.zba"


Dim objUpLoadFile
Set objUpLoadFile=New TUpLoadFile


objUpLoadFile.AutoName=False
objUpLoadFile.IsManual=True
objUpLoadFile.FileSize=0
objUpLoadFile.FileName="temp.zba"
objUpLoadFile.FullPath=ZipPathFile

If objUpLoadFile.UpLoad_Form() Then
	If objUpLoadFile.SaveFile() Then
		Call InstallApp(ZipPathFile)
	End If
End If


Response.Redirect Request.ServerVariables("HTTP_REFERER")


%>