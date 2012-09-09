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
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()
ClearGlobeCache
LoadGlobeCache

Select Case LCase(Request.QueryString("act"))
	Case "view"
		Call app_view(Request.QueryString("id"))
	Case "search"
		Call app_search()
	Case "list"
		Call app_list()
	Case Else
		Response.Write "<?xml version=""1.0"" encoding=""utf-8""?><response><err><code>-1</code><runtime>0</runtime></err></response>"
		Response.End 
End Select

Sub app_list
	Dim ArtList
	Set ArtList=New TArticleList
	ArtList.Template="CATALOGFORCLIENT"
	ArtList.html=ArtList.Template
	If ArtList.Export(Request.QueryString("page"),Request.QueryString("cate"),Request.QueryString("auth"),Request.QueryString("date"),Request.QueryString("tags"),ZC_DISPLAY_MODE_INTRO) Then
		ArtList.Build
		Response.Write Replace(Replace(Replace(ArtList.html,"<#errnumber#>",Err.Number),"<#runtime#>",RunTime),"<pagecount></pagecount>","<pagecount>0</pagecount>")
	End If

End Sub

Sub app_search()
	Dim strQuestion
	strQuestion=TransferHTML(Request.QueryString("q"),"[nohtml]")
	Dim objArticle
	Set objArticle=New TArticle
	
	objArticle.LoadCache
	
	objArticle.template="CATALOGFORCLIENT"
	
	
	Dim i
	Dim j
	Dim s
	Dim aryArticleList()
	
	Dim objRS
	Dim intPageCount
	Dim objSubArticle
	
	Dim cate
	If IsEmpty(Request.QueryString("cate"))=False Then
	cate=CInt(Request.QueryString("cate"))
	End If
	
	
	strQuestion=Trim(strQuestion)
	
	If Len(strQuestion)>0 Then
	
		strQuestion=FilterSQL(strQuestion)
	
		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn
	
		objRS.Source="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_ID]>0) AND ([log_Level]>2)"
	
		If ZC_MSSQL_ENABLE=False Then
			objRS.Source=objRS.Source & "AND( (InStr(1,LCase([log_Title]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('"&strQuestion&"'),0)<>0) )"
		Else
			objRS.Source=objRS.Source & "AND( (CHARINDEX('"&strQuestion&"',[log_Title])<>0) OR (CHARINDEX('"&strQuestion&"',[log_Intro])<>0) OR (CHARINDEX('"&strQuestion&"',[log_Content])<>0) )"
		End If
	
		If IsEmpty(cate)=False Then
			objRS.Source=objRS.Source & "AND ([log_CateID]="&cate&")"
		End If
	
		objRS.Source=objRS.Source & "ORDER BY [log_PostTime] DESC,[log_ID] DESC"
		objRS.Open()
	
		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = ZC_SEARCH_COUNT
			intPageCount=objRS.PageCount
			objRS.AbsolutePage = 1
	
			For i = 1 To objRS.PageSize
	
				ReDim Preserve aryArticleList(i)
	
				Set objSubArticle=New TArticle
				If objSubArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then
					objSubArticle.Template="PLUGINEXPORT"
					objSubArticle.SearchText=Request.QueryString("q")
					If objSubArticle.Export(ZC_DISPLAY_MODE_SEARCH)= True Then
						aryArticleList(i)=objSubArticle.subhtml
					End If
				End If
				Set objSubArticle=Nothing
	
				objRS.MoveNext
				If objRS.EOF Then Exit For
	
			Next
	
		Else
			ReDim Preserve aryArticleList(0)
		End If
	
		objRS.Close()
		Set objRS=Nothing
	
	Else
			ReDim Preserve aryArticleList(0)
	End If
	
	
	objArticle.FType=ZC_POST_TYPE_PAGE
	objArticle.Content=Join(aryArticleList)
	objArticle.Content=Replace(objArticle.Content,"<#ZC_BLOG_HOST#>",BlogHost)
	objArticle.Title=ZC_MSG085 + ":" + TransferHTML(strQuestion,"[html-format]")
	objArticle.FullRegex="{%host%}/{%alias%}.html"
	
	
	If objArticle.Export(ZC_DISPLAY_MODE_SYSTEMPAGE) Then
		objArticle.Build
		Response.Write Replace(Replace(Replace(Replace(objArticle.html,"<#articlelist/page/all#>",1),"<#articlelist/page/now#>",1),"<#errnumber#>",Err.Number),"<#runtime#>",RunTime)
	End If

End Sub

Sub app_view(ID)
	Dim Article
	Set Article=New TArticle
	If Article.LoadInfoByID(Request.QueryString("id")) Then
		If Article.Level=1 Then Call ShowError(63)
		If Article.Level=2 Then
			If Not CheckRights("Root") Then
				If (Article.AuthorID<>BlogUser.ID) Then Call ShowError(6)
			End If
		End If
		Article.Template="PLUGINEXPORT"
		If Article.Export(ZC_DISPLAY_MODE_ALL)= True Then
			Article.Build
			Response.Write Replace(Replace(Article.html,"<#errnumber#>",Err.Number),"<#runtime#>",RunTime)
		End If
	
	End If
End Sub

Call System_Terminate()

%>