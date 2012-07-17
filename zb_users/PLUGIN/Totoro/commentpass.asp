<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 

If CheckPluginState("Totoro")=False Then Call ShowError(48)

	Dim i,j
	Dim s,t
	Dim aryArticle()
	s=Request.Form("edtBatch")
	t=Split(s,",")

	ReDim Preserve aryArticle(UBound(t))
	For j=0 To UBound(t)-1
		aryArticle(j)=0
	Next

	Dim objComment
	Dim objArticle

	For i=0 To UBound(t)-1
		Set objComment=New TComment
		If objComment.LoadInfobyID(t(i)) Then
			objComment.isCheck=False
			aryArticle(i)=objComment.log_ID
			objComment.Post
		End If
		Set objComment=Nothing
	Next


	For j=0 To UBound(t)-1
		If clng(aryArticle(j))>0 Then
				Call BuildArticle(aryArticle(j),False,True)
				BlogReBuild_Comments
				Call ClearGlobeCache
				Call LoadGlobeCache
		End If
	Next


	Response.Redirect "setting1.asp"

%>
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>