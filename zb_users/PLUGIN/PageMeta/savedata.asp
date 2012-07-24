<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_manage.asp" -->
<%
Call System_Initialize()
Call GetCategory
Call GetUser
Call CheckReference("")
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("PageMeta")=False Then Call ShowError(48)
BlogTitle="PageMeta"

Dim a,b,c,d,e,f
f=Array("","Article","Category","User","Tag")
a=Request.Form("type")
b=Request.Form("id")
c=Request.Form("txaContent")
Call CheckParameter(a,"int",1)
Execute "Set d=New T" & f(a)
d.LoadInfoById b
d.Meta.SetValue "pagemeta",pAgEmEtA_EsCaPe_(c)
If a=3 Then
	Execute "Call Filter_Plugin_EditUser_Core(d)"
Else
	Execute "Call Filter_Plugin_Post"&f(a)&"_Core(d)"
End If
If a<>3 Then
	If d.Post Then
		if a=1 Then
			Call BuildArticle(d.ID,True,True)
		End If
		Execute "Call Filter_Plugin_Post"&f(a)&"_Succeed(d)"
	End If
Else
	If d.Edit(BlogUser) Then
		Execute "Call Filter_Plugin_EditUser_Succeed(d)"
	End IF
End If
Set d=nothing
Call SetBlogHint(True,True,Empty)
Response.Redirect "List.asp?act="&f(a)&"Mng"

%>
