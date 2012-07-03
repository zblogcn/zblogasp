<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)&(sipo)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_function_wap.asp
'// 开始时间:    2006-3-19
'// 最后修改:    2007-1-24
'// 备    注:    WAP函数模块
'///////////////////////////////////////////////////////////////////////////////



Function WapDelArt()
	Response.Write WapTitle(ZC_MSG063)
	call DelArticle(Request.QueryString("id"))
	Call MakeBlogReBuild_Core()
	Response.Write "<a href="""&unescape(Request.QueryString("url"))&""">"&ZC_MSG065&"</a><br/>"
	'Response.Write "<a href="""&WapLoginStr&"&amp;act=BlogReBuild"">"&ZC_MSG072&"</a><br/>"
End Function

Function WapDelCom()
	Response.Write WapTitle(ZC_MSG063)
	Call MakeBlogReBuild_Core()
	call DelComment(Request.QueryString("id"),Request.QueryString("log_id"))
	Response.Write "<a href="""&unescape(Request.QueryString("url"))&""">"&ZC_MSG065&"</a><br/>"
	'Response.Write "<a href="""&WapLoginStr&"&amp;act=BlogReBuild"">"&ZC_MSG072&"</a><br/>"
End Function

Function WapPostArt()
	Response.Write WapTitle(ZC_MSG168)
	Call MakeBlogReBuild_Core()
	call PostArticle_WAP
	'Response.Write "<a href="""&WapLoginStr&"&amp;act=BlogReBuild"">"&ZC_MSG072&"</a><br/>"
End Function


Function PostArticle_WAP()

	Dim s
	Dim objRegExp

	If Request.Form("edtID")<>"0" Then
		Dim objTestArticle
		Set objTestArticle=New TArticle
		If objTestArticle.LoadInfobyID(Request.Form("edtID")) Then
			If Not((objTestArticle.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Exit Function
		Else
			Call ShowError(9)
		End If
	End If

	Dim objArticle
	Set objArticle=New TArticle
	objArticle.ID=Request.Form("edtID")
	objArticle.CateID=Request.Form("edtCateID")
	objArticle.AuthorID=Request.Form("edtAuthorID")
	objArticle.Level=Request.Form("edtLevel")
	objArticle.PostTime=Request.Form("edtYear") & "-" & Request.Form("edtMonth") & "-" & Request.Form("edtDay") & " " &  Request.Form("edtTime")
	objArticle.Title=Request.Form("edtTitle")
	objArticle.Tag=ParseTag(Request.Form("edtTag"))
	objArticle.Alias=Request.Form("edtAlias")
	objArticle.Istop=Request.Form("edtIstop")

	objArticle.Intro=Request.Form("txaIntro")

	objArticle.Content=Request.Form("txaContent")

	If objArticle.Post Then
		Call BuildArticle(objArticle.ID,True,True)
		Call MakeBlogReBuild_Core()
		PostArticle_WAP=True
	End If

End Function


Function WapEdtArt()

	Dim Log_ID

	Response.Write WapTitle(ZC_MSG168)

	Response.Write ZC_MSG060&":<input emptyok=""false"" name=""edtTitle"" size=""15"" maxlength=""100"" value="""" /><br/>"

	Response.Write ZC_MSG012&":<select name=""edtCateID"">"
		GetCategory()
		Dim Category
		For Each Category in Categorys
			If IsObject(Category) Then
				Response.Write "<option value="""&Category.ID&""">"&TransferHTML(Category.Name,"[html-format]")&"</option>"
			End If
		Next
	Response.Write "</select><br/>"

	Response.Write ZC_MSG003&":<select name=""edtAuthorID"">"
	Response.Write "<option value="""&BlogUser.ID&""">"&TransferHTML(BlogUser.Name,"[html-format]")&"</option>"
	Response.Write "</select><br/>"

	Response.Write ZC_MSG061&"<select name=""edtLevel"">"
		Dim i
		For i=Ubound(ZVA_Article_Level_Name) to 1 step -1
			Response.Write "<option value="""& i &""">"& ZVA_Article_Level_Name(i) &"</option>"
		Next
	Response.Write "</select><br/>"

	Response.Write ZC_MSG062&":<input emptyok=""false"" name=""edtYear"" size=""10"" value="""&Year(GetTime(now()))&""" />-<input name=""edtMonth"" size=""10"" value="""&Month(GetTime(now()))&""" />-<input name=""edtDay""  size=""10"" value="""&Day(GetTime(now()))&""" />-<input name=""edtTime"" size=""12"" value="""& Hour(GetTime(now()))&":"&Minute(GetTime(now()))&":"&Second(GetTime(now()))&""" /><br/>"
	Response.Write ZC_MSG138&":<input emptyok=""true"" name=""edtTag"" size=""15"" maxlength=""100"" value="""" /><br/>"

	Response.Write ZC_MSG147&":<input emptyok=""true"" size=""10"" name=""edtAlias"" maxlength=""100"" value="""" />."&ZC_STATIC_TYPE&"<br/>"

	Response.Write ZC_MSG055&":<input emptyok=""false"" name=""txaContent"" size=""20"" maxlength=""5000"" value="""" /><br/>"
	Response.Write ZC_MSG016&":<input emptyok=""true"" name=""txaIntro"" size=""20"" maxlength=""1000"" value="""" /><br/>"
	Response.Write "<anchor>["&ZC_MSG087&"]<go href="""&WapLoginStr&"&amp;act=PostArt&amp;inpId="&Log_ID&""" method=""post"">"
	Response.Write " <postfield name=""edtID"" value=""0"" />"
	Response.Write " <postfield name=""edtTitle"" value=""$(edtTitle:n)"" />"
	Response.Write " <postfield name=""edtAuthorID"" value=""$(edtAuthorID:n)"" />"
	Response.Write " <postfield name=""edtLevel"" value=""$(edtLevel:n)"" />"
	Response.Write " <postfield name=""edtYear"" value=""$(edtYear:n)"" />"
	Response.Write " <postfield name=""edtMonth"" value=""$(edtMonth:n)"" />"
	Response.Write " <postfield name=""edtDay"" value=""$(edtDay:n)"" />"
	Response.Write " <postfield name=""edtTime"" value=""$(edtTime:n)"" />"
	Response.Write " <postfield name=""edtTag"" value=""$(edtTag:n)"" />"
	Response.Write " <postfield name=""edtAlias"" value=""$(edtAlias:n)"" />"
	Response.Write " <postfield name=""edtCateID"" value=""$(edtCateID:n)"" />"
	Response.Write " <postfield name=""txaContent"" value=""$(txaContent:n)"" />"
	Response.Write " <postfield name=""txaIntro"" value=""$(txaIntro:n)"" />"

	Response.Write "</go>"
	Response.Write "</anchor>"
End Function


Function WapCate()
	Dim Category
	Response.Write WapTitle(ZC_MSG214)
		For Each Category in Categorys
			If IsObject(Category) Then
				Response.Write Category.ID&",<a href="""&WapLoginStr&"&amp;act=Main&amp;cate="&Category.ID&""">"&TransferHTML(Category.Name,"[html-format]")&"("&Category.Count&")</a><br/>"
			End If
		Next
End Function

Function WapStat()
	Response.Write WapTitle(ZC_MSG029)
	Response.Write  Replace(Replace(LoadFromFile(BlogPath & "zb_users/INCLUDE/statistics.asp","utf-8"),"<li>",""),"</li>","<br/>")
End Function

Function WapAddCom(PostType)

	If ZC_WAPCOMMENT_ENABLE=False Then Call ShowError(40): Exit Function
	
	Dim log_ID,Author,Content,Email,HomePage
	log_ID=Request("inpId")
	Author=Request.Form("inpName")
	Content=Request.Form("inpArticle")
	Email=Request.Form("inpEmail")
	HomePage=Request.Form("inpHomePage")

	Call CheckParameter(log_ID,"int",0)
	If log_ID=0 Then Call ShowError(3): Exit Function

	Response.Write WapTitle(ZC_MSG211)

    'Response.Write "<input type=""hidden"" name=""inpId"" id=""inpId"" value="""&Log_ID&""" />"
  
    If PostType<>0 Then
    Response.Write ZVA_ErrorMsg(PostType)&"<br/>"
	End If
	
	If PostType=31 Then
	Response.Write ZC_MSG001&"*:<input type=""text"" emptyok=""false"" name=""inpName"" size=""12"" value="""&BlogUser.Name&""" maxlength="""&ZC_USERNAME_MAX&"""/><br/>"
	Else
	Response.Write ZC_MSG001&"*:<input type=""text"" emptyok=""false"" name=""inpName"" size=""12"" value="""&Author&""" maxlength="""&ZC_USERNAME_MAX&"""/><br/>"
	End If

	If PostType=6 Then
	Response.Write ZC_MSG002&":<input type=""password"" emptyok=""false"" name=""inpPass"" size=""12"" value="""" maxlength="""&ZC_PASSWORD_MAX&"""/><br/>"
	End If

	Response.Write ZC_MSG053&":<input type=""text"" emptyok=""true"" name=""inpEmail"" size=""12"" value="""&Email&""" maxlength="""&ZC_EMAIL_MAX&"""/><br/>"
	Response.Write ZC_MSG054&":<input type=""text"" emptyok=""true"" name=""inpHomePage"" size=""12"" value="""&HomePage&""" maxlength="""&ZC_HOMEPAGE_MAX&""" /><br/>"
	Response.Write ZC_MSG055&"*("&ZC_MSG055&":1000):<br/><input type=""text"" emptyok=""false"" name=""inpArticle"" size=""20"" maxlength="""&ZC_CONTENT_MAX&""" value="""&Content&"""></input><br/>"
	
	Response.Write "<anchor title=""post"">["&ZC_MSG087&"]"
	Response.Write "<go href="""&WapLoginStr&"&amp;act=PostCom&amp;inpId="&Log_ID&""" method=""post"">"
	Response.Write "<postfield name=""username"" value=""$(inpName:n)"" />"
	If PostType=6 Then
	Response.Write "<postfield name=""password"" value=""$(inpPass:n)"" />"
	End If
	Response.Write "<postfield name=""email"" value=""$(inpEmail:n)"" />"
	Response.Write "<postfield name=""url"" value=""$(inpHomePage:n)"" />"
	Response.Write "<postfield name=""content"" value=""$(inpArticle:n)"" />"
	Response.Write "</go></anchor><br/>"
	
End Function


Function WapPostCom()

	If ZC_WAPCOMMENT_ENABLE=False Then Call ShowError(40): Exit Function

	If Not IsEmpty(Request.Form("password")) Then
	Response.Cookies("password")=md5(Request.Form("password"))
	session(ZC_BLOG_CLSID&"password")=md5(Request.Form("password"))
    Response.Cookies("username")=Request.Form("username")
	session(ZC_BLOG_CLSID&"username")=Request.Form("username")
	Call WapCheckLogin
	End IF

	Dim objComment
	Dim objArticle

	Set objComment=New TComment
	Set objArticle=New TArticle

	objComment.log_ID=Request("inpID")
	objComment.AuthorID=BlogUser.ID
	objComment.Author=Request.Form("username")
	objComment.Content=Request.Form("content")
	objComment.Email=Request.Form("email")
	objComment.HomePage=Request.Form("url")

	If Not CheckRegExp(objComment.Author,"[username]") Then Call  WapAddCom(15) :Exit Function
	
	IF Len(objComment.Email)>0 Then
		If Not CheckRegExp(objComment.Email,"[email]") Then Call  WapAddCom(29) :Exit Function
	End If

	IF Len(objComment.HomePage)>0 Then
		If InStr(objComment.HomePage,"http")=0 Then objComment.HomePage="http://" & objComment.HomePage
		If Not CheckRegExp(objComment.HomePage,"[homepage]") Then Call WapAddCom(30) :Exit Function
	End If

	IF Len(objComment.Content)>ZC_CONTENT_MAX Then Call WapAddCom(46) :Exit Function

	Dim objUser
	
	For Each objUser in Users

		If IsObject(objUser) Then

		    '没有登陆
			If (UCase(objUser.Name)=UCase(objComment.Author)) And (objUser.ID<>objComment.AuthorID) Then
			Call WapAddCom(6)
			Exit Function
			End If

			'已经登陆了用不同的用户名
			If (objUser.ID=objComment.AuthorID) And (UCase(objUser.Name)<>UCase(objComment.Author)) Then
			Call WapAddCom(31)
			Exit Function
			End If

			'完全符合
			If (objUser.ID=objComment.AuthorID) And (UCase(objUser.Name)=UCase(objComment.Author)) Then	
				objComment.Author=objUser.Name
			End If

		End If

	Next
		
	Dim objRS
	Dim strSpamIP
	Dim strSpamContent

	Set objRS=objConn.Execute("SELECT [comm_IP],[comm_Content] FROM [blog_Comment] WHERE [comm_ID]= ( SELECT MAX(comm_ID) FROM [blog_Comment] )")

	If (Not objRS.bof) And (Not objRS.eof) Then
		strSpamIP=objRS("comm_IP")
		strSpamContent=objRS("comm_Content")
	End If

	objRS.Close
	Set objRS=Nothing

	If (strSpamContent=objComment.Content) Then
		Call WapAddCom(39)
		Exit Function
	End If

	If objComment.Post Then
		If objArticle.LoadInfoByID(objComment.log_ID) Then
			Call BuildArticle(objArticle.ID,False,False)
			BlogReBuild_Comments
			WapPostCom=True
		End If
	End if

	Response.Write WapTitle(ZC_MSG211)
    Response.Write "<a href="""&WapLoginStr&"&amp;act=View&amp;id="&objComment.log_ID&""">"&ZC_MSG065&ZC_MSG048&"</a><br/>"


	Set objComment=Nothing

	
End Function

Function WapLogin()
	Dim User,Password

	User=Request.Form("username")
	Password=Request.Form("password")
	Call CheckParameter(User,"sql",Empty)
	Call CheckParameter(Password,"sql",Empty)

	If IsEmpty(User) OR IsEmpty(Password) Then
	Response.Write WapTitle(ZC_MSG009)
	Response.Write ZC_MSG001&":<input type=""text"" name=""username"" size=""12"" value="""" /><br/>"
	Response.Write ZC_MSG002&":<input type=""password"" name=""password"" size=""12"" value="""" /><br/>"
	Response.Write "<anchor title=""post"">["&ZC_MSG087&"]<go href="""&ZC_FILENAME_WAP&"?act=Login"" method=""post""><postfield name=""username"" value=""$(username:n)"" /><postfield name=""password"" value=""$(password:n)"" /></go></anchor><br/>"
	Else
		Response.Cookies("password")=md5(Password)
		session(ZC_BLOG_CLSID&"password")=md5(Password)
		Response.Cookies("username")=User
		session(ZC_BLOG_CLSID&"username")=User

		If BlogUser.Verify=False Then
			Call ShowError(8)
		Else
			Response.Write WapTitle(ZC_MSG009)
		End If

	End If

End Function

Function WapMenu()

        Response.Write WapTitle(ZC_BLOG_TITLE)
		Response.Write "<a href="""&WapLoginStr&"&amp;act=Login"">"&ZC_MSG009&"</a><br/>"
		Response.Write "<a href="""&WapLoginStr&"&amp;act=Com"">"&ZC_MSG027&"</a><br/>"
		Response.Write "<a href="""&WapLoginStr&"&amp;act=Main"">"&ZC_MSG032&"</a><br/>"
		Response.Write "<a href="""&WapLoginStr&"&amp;act=Cate"">"&ZC_MSG214&"</a><br/>"
		Response.Write "<a href="""&WapLoginStr&"&amp;act=Stat"">"&ZC_MSG029&"</a><br/>"	
		If BlogUser.Level<=3 Then
		Response.Write "<a href="""&WapLoginStr&"&amp;act=AddArt"">"&ZC_MSG168&"</a><br/>"	
		'Response.Write "<a href="""&WapLoginStr&"&amp;act=BlogReBuild"">"&ZC_MSG072&"</a><br/>"
		End If

End Function

Function WapMain()

        Response.Write WapTitle(ZC_BLOG_TITLE)

		Response.Write WapExport(Request("page"),Request("cate"),Request("auth"),Request("date"),Request("tags"),ZC_DISPLAY_MODE_ALL)
		Response.Write WapExportBar(Request("page"),intPageCount,Request("cate"),Request("auth"),Request("date"),Request("tags"))

End Function

Function WapCom()

        Dim i,CurrentPage,log_ID
		
		CurrentPage=Request.QueryString("page")
		log_ID=Request.QueryString("id")
		Call CheckParameter(CurrentPage,"int",1)
		Call CheckParameter(log_ID,"int",0)
		
		Dim Article
		If log_ID<>0 Then
			Set Article=New TArticle
			If Article.LoadInfoByID(log_ID) Then
				If Article.Level=1 Then Response.Write WapTitle(ZVA_Article_Level_Name(1))&ZVA_ErrorMsg(9):Exit Function
				If Article.Level=2 Then
					If Not CheckRights("Root") Then
						If (Article.AuthorID<>BlogUser.ID) Then Response.Write WapTitle(ZVA_Article_Level_Name(2))&ZVA_ErrorMsg(6):Exit Function
					End If
				End If
			End If
		End If

		Dim objRS
		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn
		If log_ID=0 Then 
		objRS.Source="SELECT blog_Comment.* , blog_Article.log_ID, blog_Article.log_Title FROM blog_Comment INNER JOIN blog_Article ON blog_Comment.log_ID = blog_Article.log_ID ORDER BY blog_Comment.comm_PostTime DESC"
		Response.Write WapTitle(ZC_MSG027)&ZC_MSG027&"<br/><br/>"
		Else
		objRS.Source="SELECT blog_Comment.* , blog_Article.log_ID, blog_Article.log_Title FROM blog_Comment INNER JOIN blog_Article ON blog_Comment.log_ID = blog_Article.log_ID WHERE blog_Comment.log_ID="&log_ID&" ORDER BY blog_Comment.comm_PostTime DESC"
		Response.Write WapTitle(Article.Title&"-"&ZC_MSG013)&Article.Title&"<br/><br/>"
		End If
		objRS.Open()

		If (Not objRS.bof) And (Not objRS.eof) Then
		
		Dim strCTemplate,ComRecordCount
		strCTemplate=GetTemplate("TEMPLATE_WAP_ARTICLE_COMMENT")
		

			objRS.PageSize = ZC_COMMENT_COUNT_WAP
			intPageCount=objRS.PageCount
			ComRecordCount=objRS.RecordCount
			objRS.AbsolutePage = CurrentPage
			

			For i=1 To objRS.PageSize
					
					Dim objComment
					Set objComment=New TComment
					If objComment.LoadInfoByArray(Array(objRS("comm_ID"),objRS("blog_Comment.log_ID"),objRS("comm_AuthorID"),objRS("comm_Author"),objRS("comm_Content"),objRS("comm_Email"),objRS("comm_HomePage"),objRS("comm_PostTime"),objRS("comm_IP"),objRS("comm_Agent"))) Then
					Dim strC_Count
					strC_Count=ComRecordCount-((CurrentPage-1)*ZC_COMMENT_COUNT_WAP+i)+1

					ReDim Preserve aryStrC(i)
					aryStrC(i)=strCTemplate
					aryStrC(i)=Replace(aryStrC(i),"<#ZC_FILENAME_WAP#>",ZC_FILENAME_WAP)
					aryStrC(i)=Replace(aryStrC(i),"<#article/id#>",objRS("blog_Comment.log_ID"))
					aryStrC(i)=Replace(aryStrC(i),"<#article/title#>",objRS("log_Title"))
					aryStrC(i)=Replace(aryStrC(i),"<#article/comment/id#>",objComment.ID)
					aryStrC(i)=Replace(aryStrC(i),"<#article/comment/name#>",objComment.Author)
					aryStrC(i)=Replace(aryStrC(i),"<#article/comment/url#>",objComment.HomePage)
					aryStrC(i)=Replace(aryStrC(i),"<#article/comment/email#>",objComment.SafeEmail)
					aryStrC(i)=Replace(aryStrC(i),"<#article/comment/posttime#>",FormatDateTime(objComment.PostTime,vbShortDate)&" "&FormatDateTime(objComment.PostTime,vbShortTime))
					aryStrC(i)=Replace(aryStrC(i),"<#article/comment/content#>",TransferHTML(TransferHTML(UBBCode(objComment.HtmlContent,"[face][link][autolink][font][code][image][typeset][media][flash][key][upload]"),"[html-japan][vbCrlf][upload]"),"[wapnohtml]"))
					aryStrC(i)=Replace(aryStrC(i),"<#article/comment/count#>",strC_Count)
					aryStrC(i)=Replace(aryStrC(i),"<#article/comment/authorid#>",objComment.AuthorID)
					aryStrC(i)=Replace(aryStrC(i),"<#article/comment/firstcontact#>",objComment.FirstContact)
				    
					If BlogUser.Level<=3 Then
						aryStrC(i)=Replace(aryStrC(i),"<#url#>",Escape(ZC_BLOG_HOST&ZC_FILENAME_WAP&"?mode=WAP&amp;"&Replace(Request.QueryString,"&","&amp;")))
						aryStrC(i)=Replace(aryStrC(i),"<#adbegin#>","")
						aryStrC(i)=Replace(aryStrC(i),"<#adend#>","")
					Else
						Dim objRegExp
						Set objRegExp=New RegExp
						objRegExp.IgnoreCase =True
						objRegExp.Global=True
						objRegExp.Pattern="<#adbegin#>(.+)<#adend#>"
						aryStrC(i)= objRegExp.Replace(aryStrC(i),"")
					End If
				
					Dim aryTemplateTagsName,aryTemplateTagsValue

					aryTemplateTagsName=TemplateTagsName
					aryTemplateTagsValue=TemplateTagsValue


					aryTemplateTagsName(0)="BlogTitle"
					aryTemplateTagsValue(0)=ZC_BLOG_TITLE
					
					Dim k
					For k=0 to UBound(aryTemplateTagsName)
					    aryStrC(i)=Replace(aryStrC(i),"<#" & aryTemplateTagsName(k) & "#>",aryTemplateTagsValue(k))
					Next

					End If


					Set objComment=Nothing

				objRS.MoveNext
				If objRS.EOF Then Exit For

			Next

		Else

			Exit Function

		End If

		objRS.Close()
		Set objRS=Nothing
		
		Dim strC
		strC=Join(aryStrC)
		
		Dim PageBar
		PageBar="<a href="""&WapLoginStr&"&amp;act=Com&amp;id="&log_ID&"&amp;Page=1"">[&lt;&lt;]</a>"
			
		For i=CurrentPage-Cint(ZC_COMMENT_PAGEBAR_COUNT_WAP/2) to CurrentPage+Cint(ZC_COMMENT_PAGEBAR_COUNT_WAP/2) 
		
		If i>0 and i<=intPageCount Then
			If i=CurrentPage Then
			PageBar=PageBar&"<a href="""&WapLoginStr&"&amp;act=Com&amp;id="&log_ID&"&amp;Page="&i&""">[["&i&"]]</a>"
			Else
			PageBar=PageBar&"<a href="""&WapLoginStr&"&amp;act=Com&amp;id="&log_ID&"&amp;Page="&i&""">["&i&"]</a>"
			End If
		End If

		Next

		PageBar=PageBar&"<a href="""&WapLoginStr&"&amp;act=Com&amp;id="&log_ID&"&amp;Page="&intPageCount&""">[&gt;&gt;]</a>"
		
		Response.Write strC&PageBar

End Function

Function WapView()
	Dim Article,ZC_SINGLE_START,CurrentPage,i,log_ID
	CurrentPage=Request.QueryString("page")
	log_ID=Request.QueryString("id")
	Call CheckParameter(CurrentPage,"int",1)
	Call CheckParameter(log_ID,"int",0)
	
	If log_ID=0 Then Call ShowError(3) : Exit Function

	Set Article=New TArticle
	If Article.LoadInfoByID(log_ID) Then
			'If BlogUser.Level>

			If Article.Level=1 Then Response.Write WapTitle(ZVA_Article_Level_Name(1))&ZVA_ErrorMsg(9):Exit Function
			If Article.Level=2 Then
				If Not CheckRights("Root") Then
					If (Article.AuthorID<>BlogUser.ID) Then Response.Write WapTitle(ZVA_Article_Level_Name(2))&ZVA_ErrorMsg(6):Exit Function
				End If
			End If

	        Response.Write WapTitle(Article.Title)
			Dim ArticleContent,PageCount,PageBar
			ArticleContent=TransferHTML(TransferHTML(UBBCode(Article.Content,"[face][link][autolink][font][code][image][typeset][media][flash][key][upload]"),"[html-japan][vbCrlf][upload]"),"[wapnohtml]")
			
			PageCount = Int(Len(ArticleContent)/ZC_SINGLE_SIZE_WAP) + 1
			ZC_SINGLE_START=Cint((CurrentPage-1)*ZC_SINGLE_SIZE_WAP+1)
			If ZC_SINGLE_START<1 Then ZC_SINGLE_START=1
			ArticleContent=Mid(ArticleContent,ZC_SINGLE_START,ZC_SINGLE_SIZE_WAP)
			ArticleContent=TransferHTML(ArticleContent,"[html-format][wapnohtml][nbsp-br]")

		    PageBar="<br/><a href="""&WapLoginStr&"&amp;act=View&amp;id="&log_ID&"&amp;Page=1"">[&lt;&lt;]</a>"
			
			For i=CurrentPage-Cint(ZC_SINGLE_PAGEBAR_COUNT_WAP/2) to CurrentPage+Cint(ZC_SINGLE_PAGEBAR_COUNT_WAP/2)
			
			If i>0 and i<=PageCount Then
				If i=CurrentPage Then
				PageBar=PageBar&"<a href="""&WapLoginStr&"&amp;act=View&amp;id="&log_ID&"&amp;Page="&i&""">[["&i&"]]</a>"
				Else
				PageBar=PageBar&"<a href="""&WapLoginStr&"&amp;act=View&amp;id="&log_ID&"&amp;Page="&i&""">["&i&"]</a>"
				End If
			End If
			Next
			PageBar=PageBar&"<a href="""&WapLoginStr&"&amp;act=View&amp;id="&log_ID&"&amp;Page="&PageCount&""">[&gt;&gt;]</a>"
			ArticleContent=ArticleContent&PageBar


			If Article.Export(ZC_DISPLAY_MODE_ALL) Then
			    Article.template_Wap="wap_single"
				Article.Build
				Article.htmlWAP=Replace(Article.htmlWAP,"<#article/PageContent#>",ArticleContent)
				Article.htmlWAP=Replace(Article.htmlWAP,"<#ZC_FILENAME_WAP#>",ZC_FILENAME_WAP)
			    If BlogUser.Level<=3 Then
					Article.htmlWAP=Replace(Article.htmlWAP,"<#url#>",Escape(ZC_BLOG_HOST&ZC_FILENAME_WAP&"?mode=WAP&amp;"&Replace(Request.QueryString,"&","&amp;")))
					Article.htmlWAP=Replace(Article.htmlWAP,"<#adbegin#>","")
					Article.htmlWAP=Replace(Article.htmlWAP,"<#adend#>","")
				Else
					Dim objRegExp
					Set objRegExp=New RegExp
					objRegExp.IgnoreCase =True
					objRegExp.Global=True
					objRegExp.Pattern="<#adbegin#>(.+)<#adend#>"
					Article.htmlWAP= objRegExp.Replace(Article.htmlWAP,"")
				End If
				Response.Write Article.htmlWAP
			End If
			
	End If


End Function

Function WapExport(intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType)

		Dim i,j
		Dim objRS
		Dim objArticle

		Call CheckParameter(intPage,"int",1)
		Call CheckParameter(intCateId,"int",Empty)
		Call CheckParameter(intAuthorId,"int",Empty)
		Call CheckParameter(dtmYearMonth,"dtm",Empty)
		
		Dim Title
		Title=ZC_BLOG_SUBTITLE

		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn
		objRS.Source="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_IsTop] FROM [blog_Article] WHERE ([log_ID]>0) AND ([log_Level]>1)"

		If Not IsEmpty(intCateId) Then
			objRS.Source=objRS.Source & "AND([log_CateID]="&intCateId&")"
			On Error Resume Next
			Title=Categorys(intCateId).Name
			Err.Clear
		End if
		If Not IsEmpty(intAuthorId) Then
			objRS.Source=objRS.Source & "AND([log_AuthorID]="&intAuthorId&")"
			On Error Resume Next
			Title=Users(intAuthorId).Name
			Err.Clear
		End if
		If IsDate(dtmYearMonth) Then
			Dim y
			Dim m
			Dim ny
			Dim nm

			If IsDate(dtmYearMonth) Then
				dtmYearMonth=CDate(dtmYearMonth)
			Else
				Call ShowError(3)
			End If

			y=year(dtmYearMonth)
			m=month(dtmYearMonth)
			ny=y
			nm=m+1
			If m=12 Then ny=ny+1:nm=1

			objRS.Source=objRS.Source & "AND([log_PostTime] BETWEEN #"&y&"-"&m&"-1# AND #"&ny&"-"&nm&"-1#)"

			Application.Lock
			If Year(dtmYearMonth)=Year(GetTime(now())) And Month(dtmYearMonth)=Month(GetTime(now())) Then
				Template_Calendar=Application(ZC_BLOG_CLSID & "CACHE_INCLUDE_CALENDAR")
			End If
			Application.UnLock

			Title=Year(dtmYearMonth) & " " & ZVA_Month(Month(dtmYearMonth))
		End If
		If Not IsEmpty(strTagsName) Then
			On Error Resume Next
			Dim Tag
			For Each Tag in Tags
				If IsObject(Tag) Then
					If UCase(Tag.Name)=UCase(strTagsName) Then
						objRS.Source=objRS.Source & "AND([log_Tag] LIKE '%{" & Tag.ID & "}%')"
					End If
				End If
			Next
			Err.Clear
			Title=strTagsName
		End If

		objRS.Source=objRS.Source & "ORDER BY [log_PostTime] DESC,[log_ID] DESC"
		objRS.Open()

		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = ZC_DISPLAY_COUNT_WAP
			intPageCount=objRS.PageCount
			objRS.AbsolutePage = intPage

			For i = 1 To objRS.PageSize

				ReDim Preserve aryArticleList(i)

				Set objArticle=New TArticle
				If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16))) Then
					If objArticle.Export(intType)= True Then
						aryArticleList(i)=objArticle.Template_Article_Multi_WAP
					End If
				End If
				Set objArticle=Nothing

				objRS.MoveNext
				If objRS.EOF Then Exit For

			Next

		Else

			Exit Function

		End If

		objRS.Close()
		Set objRS=Nothing
			
			Dim Template_Article_Multi
			Template_Article_Multi=Join(aryArticleList)

		Dim Template_Calendar
		If IsEmpty(Template_Calendar) Or Len(Template_Calendar)=0 Then
			Application.Lock
			Template_Calendar=Application(ZC_BLOG_CLSID & "CACHE_INCLUDE_CALENDAR")
			Application.UnLock
		End If
		

		Dim aryTemplateTagsName,aryTemplateTagsValue

		aryTemplateTagsName=TemplateTagsName
		aryTemplateTagsValue=TemplateTagsValue

		aryTemplateTagsName(0)="BlogTitle"
		aryTemplateTagsValue(0)=Title

		j=UBound(aryTemplateTagsName)

		For i=0 to j
			Template_Article_Multi=Replace(Template_Article_Multi,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
		Next

		Template_Article_Multi=Replace(Template_Article_Multi,"<#ZC_FILENAME_WAP#>",ZC_FILENAME_WAP)
		If BlogUser.Level<=3 Then
			Template_Article_Multi=Replace(Template_Article_Multi,"<#url#>",Escape(ZC_BLOG_HOST&ZC_FILENAME_WAP&"?mode=WAP&amp;"&Replace(Request.QueryString,"&","&amp;")))

			Template_Article_Multi=Replace(Template_Article_Multi,"<#adbegin#>","")
			Template_Article_Multi=Replace(Template_Article_Multi,"<#adend#>","")
		Else
			Dim objRegExp
			Set objRegExp=New RegExp
			objRegExp.IgnoreCase =True
			objRegExp.Global=True
			objRegExp.Pattern="<#adbegin#>(.+)<#adend#>"
			Template_Article_Multi= objRegExp.Replace(Template_Article_Multi,"")
		End If


		WapExport=Template_Article_Multi

	End Function

	Function WapExportBar(intNowPage,intAllPage,intCateId,intAuthorId,dtmYearMonth,strTagsName)

		Dim i
		Dim s
		Dim t
		Dim strPageBar



		If Not IsEmpty(intCateId) Then t=t & "cate=" & intCateId & "&amp;"
		If Not IsEmpty(dtmYearMonth) Then t=t & "date=" & Year(dtmYearMonth) & "-" & Month(dtmYearMonth) & "&amp;"
		If Not IsEmpty(intAuthorId) Then t=t & "auth=" & intAuthorId & "&amp;"
		If Not (strTagsName="") Then t=t & "tags=" & Server.URLEncode(strTagsName) & "&amp;"
		If intAllPage>0 Then
			Dim a,b

			s=""&WapLoginStr&"&amp;act=Main&amp;"& t &"page=1"
		
			strPageBar="<a href=""<#pagebar/page/url#>"">[<#pagebar/page/number#>]</a>"

			strPageBar=Replace(strPageBar,"<#pagebar/page/url#>",s)
			strPageBar=Replace(strPageBar,"<#pagebar/page/number#>","&lt;&lt;")
			Dim Template_PageBar
			Template_PageBar=Template_PageBar & strPageBar

			If intAllPage>ZC_PAGEBAR_COUNT_WAP Then
				a=intNowPage
				b=intNowPage+ZC_PAGEBAR_COUNT_WAP
				If a>ZC_PAGEBAR_COUNT_WAP Then a=a-1:b=b-1
				If b>intAllPage Then b=intAllPage:a=intAllPage-ZC_PAGEBAR_COUNT_WAP
			Else
				a=1:b=intAllPage
			End If
			For i=a to b
				If i>0 Then

					s=""&WapLoginStr&"&amp;act=Main&amp;"& t &"page="& i

					strPageBar="<a href=""<#pagebar/page/url#>"">[<#pagebar/page/number#>]</a>"
					strPageBar=Replace(strPageBar,"<#pagebar/page/url#>",s)
					strPageBar=Replace(strPageBar,"<#pagebar/page/number#>",i)
					Template_PageBar=Template_PageBar & strPageBar
				End If
			Next

			s=""&WapLoginStr&"&amp;act=Main&amp;"& t &"page="& intAllPage

			strPageBar="<a href=""<#pagebar/page/url#>"">[<#pagebar/page/number#>]</a>"
			strPageBar=Replace(strPageBar,"<#pagebar/page/url#>",s)
			strPageBar=Replace(strPageBar,"<#pagebar/page/number#>","&gt;&gt;")
			Template_PageBar=Template_PageBar & strPageBar
			
			Dim Template_PageBar_Previous
			If intNowPage=1 Then
				Template_PageBar_Previous=""
			Else
				Template_PageBar_Previous="<a href="""&WapLoginStr&"&amp;act=Main&amp;"& t &"page="& intNowPage-1 &""">"&ZC_MSG156&"</a>"

			End If

			Dim Template_PageBar_Next
			If intNowPage=intAllPage Then
				Template_PageBar_Next=""
			Else
				Template_PageBar_Next="<a href="""&WapLoginStr&"&amp;act=Main&amp;"& t &"page="& intNowPage+1 &""">"&ZC_MSG155&"</a>"
			End If

		End If

		WapExportBar=Template_PageBar

	End Function

	Public Function WapTitle(Str)
	WapTitle="<card title="""&Str&""" id=""card1"">"&vbnewline&"<p>"&WapCheckLogin&"</p><p>"&vbnewline
	End Function

	Public Function WapError()
		dim ID
		ID=Request("id")
		If Not IsNumeric(ID) Then
		ID=0
		ElseIf CINT(ID)>Ubound(ZVA_ErrorMsg) Or CINT(ID)<0 Then
		ID=0
		End If
		Response.Write WapTitle(ZVA_ErrorMsg(ID))&ZVA_ErrorMsg(ID)
	End Function

   Function WapLoginStr()
	   WapLoginStr=ZC_BLOG_HOST&ZC_FILENAME_WAP&"?mode=WAP"
	End Function
 
 Function WapCheckLogin()
	Dim username,password,s
		username=Request.Form("username")
		 password=Request.Form("password")
	   If (Not IsEmpty(Request.Cookies("username"))) And (Not IsEmpty(Request.Cookies("password"))) Then
	   username=Request.Cookies("username")
	   password=Request.Cookies("password")
	   session(ZC_BLOG_CLSID&"username")=username
	   session(ZC_BLOG_CLSID&"password")=password
	   ElseIf (Not IsEmpty(session(ZC_BLOG_CLSID&"username"))) And (Not IsEmpty(session(ZC_BLOG_CLSID&"password"))) Then
	   username=session(ZC_BLOG_CLSID&"username")
	   password=session(ZC_BLOG_CLSID&"password")
	   Request.Cookies("username")=username
	   Request.Cookies("password")=password
	   End If

	   BlogUser.LoginType="Self"
	   BlogUser.Password=password
	   BlogUser.Name=username
       BlogUser.Verify

	   
	   
	    s=BlogUser.Name&" "&ZVA_User_Level_Name(BlogUser.Level)&""
		If BlogUser.ID<>0 Then
		s=s&" <a href="""&WapLoginStr&"&amp;act=Logout&amp;url="&Escape(ZC_BLOG_HOST&ZC_FILENAME_WAP&"?mode=WAP&amp;"&Replace(Request.QueryString,"&","&amp;"))&""">"&ZC_MSG020&"</a><br/><br/>"
		End If

		WapCheckLogin=s
	   
 End Function

 Function WapCopyRight()

	 WapCopyRight=vbsunescape(Request.Cookies("username"))
	
 End Function

 Function WapLogout()

    Response.Write WapTitle(ZC_MSG020)

	Response.Cookies("username")=""
 	Response.Cookies("password")=""
	session(ZC_BLOG_CLSID&"password")=""    
	session(ZC_BLOG_CLSID&"username")=""
	Response.Cookies("username")=Empty
 	Response.Cookies("password")=Empty
	session(ZC_BLOG_CLSID&"password")=Empty    
	session(ZC_BLOG_CLSID&"username")=Empty

    Response.Write "<a href="""&unescape(Request.QueryString("url"))&""">"&ZC_MSG065&"</a><br/>"

 End Function


 Function ShowError_WAP(id)
 	Response.Redirect ZC_BLOG_HOST&ZC_FILENAME_WAP&"?act=Err&id="&id
 End Function

%>
