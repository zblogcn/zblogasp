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
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("AdvancedFunction")=False Then Call ShowError(48)
BlogTitle="增强侧栏"
Dim subCate
%>
<%

js()
Dim attr2
Select Case Request.QueryString("act")
	Case "del"
		advancedfunction.cls.config.Remove "分类_"&Request.QueryString("id")
	Case "save"
		For Each attr2 In Request.Form
			If Left(attr2,2)="m_" Then
				advancedfunction.cls.config.Write Right(attr2,Len(attr2)-2),Request.Form(attr2)
			Else
				Select Case attr2
					Case "newCategory"
						advancedfunction.cls.config.Write "分类_"&Request.Form(attr2),10
				End Select
			End If	
			
			'Response.Write attr2
		Next
End Select
advancedfunction.cls.config.Save
SetBlogHint True,Empty,Empty
Response.Redirect "main.asp"
%>
<script language="javascript" runat="server">
function js(){
	advancedfunction.init();
}
</script>

