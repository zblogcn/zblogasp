<%@ CODEPAGE=65001 %>
<% Option Explicit %>
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
	
	If s<>"" Then
	
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
				objComment.Del
			End If
			Set objComment=Nothing
		Next


		For j=0 To UBound(t)-1
			If aryArticle(j)>0 Then
				'Call BuildArticle(aryArticle(j),False,False)
			End If
		Next

		BlogReBuild_Comments
		BlogReBuild_GuestComments
	
	ElseIf request.QueryString("act")="delALL" Then

		Dim strSQL
		if ZC_MSSQL_ENABLE then
			strSQL="WHERE ([comm_isCheck]=1) "
		else
			strsql="WHERE ([comm_isCheck]=FALSE)"
		end if
		If CheckRights("Root")=False Then strSQL=strSQL & "AND( (SELECT [log_AuthorID] FROM [blog_Article] WHERE [blog_Article].[log_ID] =[blog_Comment].[log_ID] ) =" & BlogUser.ID & ")"
		
		objConn.Execute("DELETE FROM [blog_Comment] " & strSQL)
		
		BlogReBuild_Comments
		BlogReBuild_GuestComments
	
	End If

	Response.Redirect "setting1.asp"

%>
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>