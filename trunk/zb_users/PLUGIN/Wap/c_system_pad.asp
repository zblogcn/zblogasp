<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)&(sipo)&(月上之木)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_system_wap.asp
'// 开始时间:    2006-03-19
'// 最后修改:    2011-08-03
'// 备    注:    WAP函数模块
'///////////////////////////////////////////////////////////////////////////////





Class TPad

	Public Title

	Public html
	Public subhtml
	Public subhtml_TemplateName

	Public ListType

	Private Ftemplate
	Public Property Let Template(strFileName)
		Ftemplate=GetTemplate("TEMPLATE_" & strFileName)
	End Property
	Public Property Get Template
		If Ftemplate="" Then
			Ftemplate=GetTemplate("TEMPLATE_DEFAULT")
		End If

		Template = Ftemplate
	End Property


	Private Ffullregex
	Public Property Let FullRegex(s)
		Ffullregex=s
	End Property
	Public Property Get FullRegex
		If Ffullregex<>"" Then 
			FullRegex=Ffullregex
		Else
			FullRegex=ZC_DEFAULT_REGEX
		End If
	End Property


	Public Url
	Private MixUrl


	Public Property Get HtmlTitle
		HtmlTitle=TransferHTML(Title,"[html-japan][html-format]")
	End Property


	Public Function Export(intPage,anyCate,anyAuthor,dtmDate,anyTag,intType)

		Title=ZC_BLOG_SUBTITLE

		Dim ArtList
		Set ArtList=New TArticleList
		ArtList.html=Template

		If ArtList.Export(intPage,anyCate,anyAuthor,dtmDate,anyTag,intType) Then
			html=ArtList.html
			Title=ArtList.Title
		End If

		Url=Replace(Replace(Url,"//","/"),":/","://",1,1)

		Export=True


	End Function



	Public Function View(intId)

		Dim Article
		Set Article=New TArticle
		Article.Template="SINGLE"

		If Article.LoadInfoByID(intId) Then

			If Article.Level=1 Then Call ShowError(63)
			If Article.Level=2 Then
				If CheckRights("Root")=False And CheckRights("ArticleAll")=False Then
					If (Article.AuthorID<>BlogUser.ID) Then Call ShowError(6)
				End If
			End If

			If Article.Export(ZC_DISPLAY_MODE_ALL)= True Then
				html=Article.html
				Title=Article.Title
			End If
		Else

		End If

	End Function


	Function Search(q)


		Dim strQuestion
		strQuestion=TransferHTML(q,"[nohtml]")

		Dim objArticle
		Set objArticle=New TArticle

		Dim i
		Dim j
		Dim s
		Dim aryArticleList()

		Dim objRS
		Dim intPageCount
		Dim objSubArticle

		Dim cate
		If IsEmpty(Request.QueryString("cate"))=False Then
		cate=CLng(Request.QueryString("cate"))
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

		objArticle.Template="PAGE"
		objArticle.FType=ZC_POST_TYPE_PAGE
		objArticle.Content=Join(aryArticleList)
		objArticle.Content=Replace(objArticle.Content,"<#ZC_BLOG_HOST#>",BlogHost)
		objArticle.Title=ZC_MSG085 + ":" + TransferHTML(strQuestion,"[html-format]")
		objArticle.FullRegex="{%host%}/{%alias%}.html"

		If objArticle.Export(ZC_DISPLAY_MODE_SYSTEMPAGE) Then
			'objArticle.Build
			'Response.Write objArticle.html
			html=objArticle.html
		End If

		Title=objArticle.Title
		Call SetVar("PAD_SIDE","")
		Call SetVar("PAD_AUTOSCREEN","")
	End Function


	Public Function Build()

		Call SetVar("PAD_NAV",FunNav())

		Call SetVar("PAD_SIDE",FunCatalogs & FunSearch() & FunAdmin())

		Call SetVar("COOKIESPATH",CookiesPath())

		Call SetVar("PAD_AUTOSCREEN","autoscreen();")

		Call SetVar("PAD_FIRSTNAME",BlogUser.FirstName)
		Call SetVar("PAD_NAME",BlogUser.Name)
		Call SetVar("PAD_EMAIL",BlogUser.EMail)
		Call SetVar("PAD_HOMEPAGE",BlogUser.HomePage)	
		
		Dim i,j

		Dim aryTemplateTagsName
		Dim aryTemplateTagsValue

		TemplateTagsDic.Item("BlogTitle")=HtmlTitle
		TemplateTagsDic.Item("ZC_BLOG_HOST")=BlogHost

		aryTemplateTagsName=TemplateTagsDic.Keys
		aryTemplateTagsValue=TemplateTagsDic.Items

		j=UBound(aryTemplateTagsName)
		For i=1 to j
			html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
		Next
		html=Replace(html,"<#" & aryTemplateTagsName(0) & "#>",aryTemplateTagsValue(0))

		Build=True

	End Function


	Function Comment()

		Dim objArticle
		Set objArticle=New TArticle
		If objArticle.LoadInfoByID(Request.Form("inpID")) Then
			Call PostComment(objArticle.CommentKey,CLng(Request.Form("inpRevID")))
		Else

		End If

	End Function



Public Function Errors(id)
	If Not IsNumeric(ID) Then
		ID=0
	ElseIf CLng(ID)>Ubound(ZVA_ErrorMsg) Or CLng(ID)<0 Then
		ID=0
	End If
	Dim s
	s=s&"<p>错误:"&ZVA_ErrorMsg(ID)&"</p>"
	s=s&"<p><span class=""stamp""><a href=""javascript:history.go(-1)"">"&ZC_MSG065&"</a></span></p>"

	Template="PAD"
	html=Template
	Call SetVar("PAD_SIDE","")
	Call SetVar("PAD_AUTOSCREEN","")
	Call SetVar("PAD_MAIN",s)
	Title="错误"

End Function


Function Login()

	Template="PAD"
	html=Template
	Call SetVar("PAD_SIDE","")
	Call SetVar("PAD_AUTOSCREEN","")

	Dim s

	s="<form id=""login"" method=""post"" action=""<#ZC_BLOG_HOST#>?mod=pad&amp;act=logging""><dl><dt>用户登录</dt><dd>用户:<input type=""text"" name=""username"" id=""username"" value="""" /></dd><dd>密码:<input type=""password"" name=""password"" id=""password"" value="""" /></dd><dd><input type=""submit"" value=""登录"" /></dd></dl></form>"

	Call SetVar("PAD_MAIN",s)
	Title="用户登录"

End Function


Function Logging()

	BlogUser.LoginType="Self"
	BlogUser.Name=Request.Form("username")
	BlogUser.PassWord=BlogUser.GetPasswordByOriginal(Request.Form("password"))

	If BlogUser.Verify=True Then

		Response.Cookies("password")=BlogUser.PassWord
		If Request.Form("savedate")<>0 Then
			Response.Cookies("password").Expires = DateAdd("d", 30, now)
		End If
		Response.Cookies("password").Path = CookiesPath()

	End If

	Response.Cookies("username")=escape(Request.Form("username"))
	If Request.Form("savedate")<>0 Then
		Response.Cookies("username").Expires = DateAdd("d", 30, now)
	End If
	Response.Cookies("username").Path = CookiesPath()

	Response.Redirect BlogHost & "?mod=pad"

End Function


Function Logout()

	Response.Cookies("username")=""
	Response.Cookies("password")=""
	Response.Cookies("username").Path = CookiesPath()
	Response.Cookies("password").Path = CookiesPath()

	Response.Redirect BlogHost & "?mod=pad"

End Function


Function FunNav()

	If BlogUser.ID>0 Then
		FunNav="<li><a href=""<#ZC_BLOG_HOST#>?mod=pad"">首页</a></li>"
	Else
		FunNav="<li><a href=""<#ZC_BLOG_HOST#>?mod=pad"">首页</a></li><li><a href=""<#ZC_BLOG_HOST#>?mod=pad&act=login"">登录</a></li>"
	End If

End Function


Function FunAdmin()

	If BlogUser.ID=0 Then Exit Function

	Dim f,s
	Set f = New TFunction
	f.Name=BlogUser.FirstName & "您好"
	f.Ftype="ul"
	s="<li><a href=""<#ZC_BLOG_HOST#>?mod=pad&amp;act=logout"">退出登录</a></li>"

	FunAdmin=Replace(f.MakeTemplate(GetTemplate("TEMPLATE_B_FUNCTION")),"<#CACHE_INCLUDE_#>",s)

End Function

Function FunSearch()

	Dim f,s
	Set f = New TFunction
	f.Name="搜索"
	f.Ftype="div"
	s="<form method=""get"" id=""search"" action=""<#ZC_BLOG_HOST#>""><input type=""hidden"" name=""mod"" value=""pad"" /><input type=""hidden"" name=""act"" value=""search"" /><input type=""text"" name=""q"" size=""9"" />&nbsp;<input type=""submit"" value=""搜"" /></form>"

	FunSearch=Replace(f.MakeTemplate(GetTemplate("TEMPLATE_B_FUNCTION")),"<#CACHE_INCLUDE_#>",s)

End Function

Function FunCatalogs()


	Call GetCategory()
	
	Dim objRS
	Dim objStream

	Dim ArtList

	'Catalogs
	Dim strCatalog,bolHasSubCate

	Dim aryCateInOrder 
	aryCateInOrder=GetCategoryOrder()


	Categorys(0).Count=CLng(objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE [log_Level]>1 AND [log_Type]=0 AND [log_CateID]=0")(0))
	If Categorys(0).Count>0 Then
		strCatalog=strCatalog & "<li class=""li-cate cate-"& Categorys(0).id &"""><a href="""& Categorys(0).HtmlUrl & """>"+Categorys(0).Name + "<span class=""article-nums""> (" & Categorys(0).Count & ")</span>" +"</a></li>"
	End If

	Dim i,j,c
	Dim strPrecata,strSubcate
	For i=Lbound(aryCateInOrder)+1 To Ubound(aryCateInOrder)
		strPrecata="":strSubcate=""
		If Categorys(aryCateInOrder(i)).ParentID=0 Then
			c=Categorys(aryCateInOrder(i)).Count

			bolHasSubCate=False
			For j=Lbound(aryCateInOrder)+1 To UBound(aryCateInOrder)
				If Categorys(aryCateInOrder(j)).ParentID=Categorys(aryCateInOrder(i)).ID Then bolHasSubCate=True
			Next
			'If bolHasSubCate Then strSubcate = "<ul class=""ul-subcates"">"
			For j=Lbound(aryCateInOrder)+1 To UBound(aryCateInOrder)
				If Categorys(aryCateInOrder(j)).ParentID=Categorys(aryCateInOrder(i)).ID And Categorys(aryCateInOrder(j)).Count>0 Then
					strSubcate=strSubcate & "<li class=""li-subcate cate-"& Categorys(aryCateInOrder(j)).ID &"""><a href="""& Categorys(aryCateInOrder(j)).Url & """>"+Categorys(aryCateInOrder(j)).Name + "<span class=""article-nums""> (" & Categorys(aryCateInOrder(j)).Count & ")</span>" +"</a></li>"
					c=c+Categorys(aryCateInOrder(j)).Count
				End If
			Next
			If bolHasSubCate And strSubcate<>"" Then strSubcate="<ul class=""ul-subcates"">" & strSubcate & "</ul>"

			If c>0 Then strPrecata="<li class=""li-cate cate-"& Categorys(aryCateInOrder(i)).ID &"""><a href="""& Categorys(aryCateInOrder(i)).Url & """>"+Categorys(aryCateInOrder(i)).Name + "<span class=""article-nums""> (" & c & ")</span>" +"</a>"& strSubcate &"</li>"

			strCatalog=strCatalog & strPrecata
		End If

	Next

	strCatalog=TransferHTML(strCatalog,"[no-asp]")

	Dim f
	Set f = New TFunction
	f.Name="所有分类"
	f.Ftype="ul"

	FunCatalogs=Replace(f.MakeTemplate(GetTemplate("TEMPLATE_B_FUNCTION")),"<#CACHE_INCLUDE_#>",strCatalog)

End Function

	Function SetVar(TemplateTag,TemplateValue)

		If IsEmpty(html) Then html=Template

		html=Replace(html,"<#" & TemplateTag & "#>",TemplateValue)

	End Function



	Function Run()

		Select Case Request.QueryString("act")
			Case "view"
				Call View(Request("id"))
			Case "Com"
				Call WapCom()
			Case "err"
				Call Errors(Request.QueryString("id"))
			Case "AddCom"		
				Call WapAddCom(0)
			Case "PostCom"
				Call WapPostCom()
			Case "DelCom"
				Call WapDelCom()
			Case "AddArt"
				Call WapEdtArt()
			Case "EdtArt"
				Call WapEdtArt()		
			Case "PostArt"
				Call WapPostArt()
			Case "DelArt"
				Call WapDelArt()
			Case "search"
				Call Search(Request("q"))
			Case "cmt"
				Call Comment()
			Case "login"
				Call Login()
			Case "logout"
				Call Logout()
			Case "logging"
				Call Logging()
			Case Else
				Call Export(Request("page"),Request("cate"),Request("auth"),Request("date"),Request("tags"),ZC_DISPLAY_MODE_ALL)		
		End Select

		Build()

		Response.Write html

	End Function


	Function ShowError(id)
		Response.Redirect BlogHost&"?mod=pad&act=err&id="&id
	End Function


	Private Sub Class_Initialize()

		ShowError_Custom="Call Pad.ShowError(id)"

		If ZC_DISPLAY_COUNT_WAP=0 Then ZC_DISPLAY_COUNT_WAP=5

		ZC_PAGEBAR_COUNT=5
		ZC_COMMENT_VERIFY_ENABLE=False

		Dim s
		s=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad.html","utf-8")

		If TemplateDic.Exists("TEMPLATE_PAD")=False Then Call TemplateDic.add("TEMPLATE_PAD","")
		TemplateDic.Item("TEMPLATE_PAD")=Replace(s,"<#PAD_MAIN#>","<#PAD_MAIN#>")

		TemplateDic.Item("TEMPLATE_DEFAULT")=Replace(s,"<#PAD_MAIN#>","<#template:article-multi#><section class=""pagebar""><#template:pagebar#></section>")
		TemplateDic.Item("TEMPLATE_SINGLE")=Replace(s,"<#PAD_MAIN#>","<#template:article-single#>")
		TemplateDic.Item("TEMPLATE_PAGE")=Replace(s,"<#PAD_MAIN#>","<#template:article-page#>")

		TemplateDic.Item("TEMPLATE_B_ARTICLE-ISTOP")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article-istop.html","utf-8")
		TemplateDic.Item("TEMPLATE_B_ARTICLE-MULTI")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article-multi.html","utf-8")
		TemplateDic.Item("TEMPLATE_B_ARTICLE-SINGLE")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article-single.html","utf-8")
		TemplateDic.Item("TEMPLATE_B_ARTICLE-PAGE")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article-page.html","utf-8")
		TemplateDic.Item("TEMPLATE_B_ARTICLE_COMMENT")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article_comment.html","utf-8")
		TemplateDic.Item("TEMPLATE_B_ARTICLE_COMMENTPOST-VERIFY")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article_commentpost-verify.html","utf-8")
		If BlogUser.ID>0 Then
			TemplateDic.Item("TEMPLATE_B_ARTICLE_COMMENTPOST")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article_commentpost2.html","utf-8")
		Else
			TemplateDic.Item("TEMPLATE_B_ARTICLE_COMMENTPOST")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article_commentpost.html","utf-8")
		End If
		TemplateDic.Item("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article_comment_pagebar.html","utf-8")
		TemplateDic.Item("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR_L")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article_comment_pagebar_l.html","utf-8")
		TemplateDic.Item("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR_R")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_article_comment_pagebar_r.html","utf-8")
		TemplateDic.Item("TEMPLATE_B_FUNCTION")=LoadFromFile(BlogPath &"zb_users\plugin\wap\template\pad_function.html","utf-8")

		ZC_POST_STATIC_MODE="ACTIVE"

		ZC_STATIC_MODE="ACTIVE"

		ZC_ARTICLE_REGEX="{%host%}/?mod=pad&act=view&id={%id%}"

		ZC_PAGE_REGEX="{%host%}/?mod=pad&act=view&id={%id%}"

		ZC_PAGE_AND_ARTICLE_PRIVATE_REGEX="{%host%}/?mod=pad&act=view&id={%id%}"

		ZC_PAGE_AND_ARTICLE_DRAFT_REGEX="{%host%}/?mod=pad&act=view&id={%id%}"

		ZC_CATEGORY_REGEX="{%host%}/?mod=pad&cate={%id%}"

		ZC_USER_REGEX="{%host%}/?mod=pad&user={%id%}"

		ZC_TAGS_REGEX="{%host%}/?mod=pad&tags={%alias%}"

		ZC_DATE_REGEX="{%host%}/?mod=pad&date={%date%}"

		ZC_DEFAULT_REGEX="{%host%}/?mod=pad"

	End Sub

End Class
'*********************************************************





%>
