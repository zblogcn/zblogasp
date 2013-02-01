<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    view.asp
'// 开始时间:    2004.07.30
'// 最后修改:    
'// 备    注:    查看页
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
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%
Dim html

Call System_Initialize()

'plugin node
For Each sAction_Plugin_View_Begin in Action_Plugin_View_Begin
	If Not IsEmpty(sAction_Plugin_View_Begin) Then Call Execute(sAction_Plugin_View_Begin)
Next

Dim objRS
Dim Article
Set Article=New TArticle

'nvap
If IsEmpty(Request.QueryString("nav"))=False Then

	If Article.LoadInfoByID(Request.QueryString("nav")) Then
		Set objRS=objConn.Execute("SELECT TOP 1 [log_FullUrl] FROM [blog_Article] WHERE ([log_ID]="& Request.QueryString("nav") &")")
		If (Not objRS.bof) And (Not objRS.eof) Then
			Response.Redirect Article.Url
		Else
			Response.Redirect BlogHost
		End If
	End If

End If



'nvap
If IsEmpty(Request.QueryString("navp"))=False Then

	If Article.LoadInfoByID(Request.QueryString("navp")) Then
		Set objRS=objConn.Execute("SELECT TOP 1 [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND ([log_Type]=0) AND ([log_PostTime]<" & ZC_SQL_POUND_KEY & Article.PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] DESC")
		If (Not objRS.bof) And (Not objRS.eof) Then
			Dim a
			Set a=New TArticle
			If a.LoadInfoByID(objRS("log_ID")) Then
				Response.Redirect a.Url
			Else
				Response.Redirect BlogHost
			End If
		Else
			Response.Redirect BlogHost
		End If
	End If

End If

'nvan
If IsEmpty(Request.QueryString("navn"))=False Then

	If Article.LoadInfoByID(Request.QueryString("navn")) Then
		Set objRS=objConn.Execute("SELECT TOP 1 [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND ([log_Type]=0) AND ([log_PostTime]>" & ZC_SQL_POUND_KEY & Article.PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] ASC")
		If (Not objRS.bof) And (Not objRS.eof) Then
			Dim b
			Set b=New TArticle
			If b.LoadInfoByID(objRS("log_ID")) Then
				Response.Redirect b.Url
			Else
				Response.Redirect BlogHost
			End If
		Else
			Response.Redirect BlogHost
		End If
	End If

End If

Dim c,d
c=Request.QueryString("id")

Set objRS=objConn.Execute("SELECT [log_ID] FROM [blog_Article] WHERE [log_Url]='"&FilterSQL(c)&"'")
If (Not objRS.bof) And (Not objRS.eof) Then
	c=objRS("log_ID")
Else

	If ZC_POST_STATIC_MODE="REWRITE" Then

		Dim fso, TxtFile
		Set fso = CreateObject("Scripting.FileSystemObject")

		If Left(d,3)="zb_" Then
			Response.Status="404 Not Found"
			Response.End
		End If
		d=c & "." & ZC_STATIC_TYPE
		If fso.FileExists(Server.MapPath(d)) Then
			Response.Write LoadFromFile(BlogPath & d,"utf-8")
			Response.End
		End If

		d=c & "/default." & ZC_STATIC_TYPE
		If fso.FileExists(Server.MapPath(d)) Then
			Response.Write LoadFromFile(BlogPath & d,"utf-8")
			Response.End
		End If

		d=c & "/index." & ZC_STATIC_TYPE
		If fso.FileExists(Server.MapPath(d)) Then
			Response.Write LoadFromFile(BlogPath & d,"utf-8")
			Response.End
		End If

	End If
End If

If TryToNumeric(c)=0 Then
	Response.Status="404 Not Found"
	Response.End
End If

If Article.LoadInfoByID(c) Then

	If Article.Level=1 Then Call ShowError(63)
	If Article.Level=2 Then
		If CheckRights("Root")=False And CheckRights("ArticleAll")=False Then
			If (Article.AuthorID<>BlogUser.ID) Then Call ShowError(6)
		End If
	End If

	If Article.Export(ZC_DISPLAY_MODE_ALL)= True Then
		If ZC_HTTP_LASTMODIFIED=True Then
			Dim strLastModified
			strLastModified=Article.PostTime
			Set objRs=objConn.Execute("SELECT TOP 1 [comm_PostTime] FROM [blog_Comment] WHERE [comm_isCheck]=0 And [log_id]=" & Article.ID)
			If Not(objRs.Eof) And Not(objRs.Bof) Then
				If DateDiff("s",objRs("comm_PostTime"),strLastModified)<0 Then strLastModified=objRs("comm_PostTime")
			End If
			Response.AddHeader "Last-Modified",ParseDateForRFC822GMT(strLastModified)
		End If
		Article.Build
		html=Article.html
		Response.Write html
	End If
Else
	Response.Status="404 Not Found"
	Response.End
End If

'plugin node
For Each sAction_Plugin_View_End in Action_Plugin_View_End
	If Not IsEmpty(sAction_Plugin_View_End) Then Call Execute(sAction_Plugin_View_End)
Next

Call System_Terminate()



Function TryToNumeric(str)
	On Error Resume Next
	Dim intN
	intN=CLng(str)
	TryToNumeric=IIf(Err.Number=0,intN,0) 
	Err.Clear
End Function
%>
<!-- <%=RunTime()%>ms --><%
If Err.Number<>0 then
	'Call ShowError(0)
End If
%>