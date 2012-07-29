<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call ZBQQConnect_Initialize()
Call CheckReference("")
If CheckPluginState("ZBQQConnect")=False Then Call ShowError(48)
If BlogUser.Level>1 Then Response.End
Dim b,a
Set a=New TConfig
a.Load "ZBQQConnect"
For b=97 To 105
	s
Next
a.Write "AppID",Request.Form("AppID")
a.Write "KEY",Request.Form("Key")
a.Save
Call SetBlogHint(True,True,Empty)
Response.Redirect "setting.asp"
Sub s()
	a.Write Chr(b),c
End Sub
Function c()
	c=IIf(Request.Form(Chr(b))="on",True,False)
End Function
%>