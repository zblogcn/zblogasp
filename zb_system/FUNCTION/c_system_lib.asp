<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:
'// 程序版本:
'// 单元名称:    c_system_lib.asp
'// 开始时间:    2004.07.25
'// 最后修改:    2007-1-4
'// 备    注:    库模块
'///////////////////////////////////////////////////////////////////////////////




'*********************************************************
' 目的：    定义TCategory类
' 输入：    无
' 返回：    无
'*********************************************************
Class TCategory

	Public ID
	Public Name
	Public Intro
	Public Order
	Public Count
	Public Alias
	Public ParentID
	Public Meta
	Public TemplateName

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property

	Public html

	Public Property Get Url

		'plugin node
		bAction_Plugin_TCategory_Url=False
		For Each sAction_Plugin_TCategory_Url in Action_Plugin_TCategory_Url
			If Not IsEmpty(sAction_Plugin_TCategory_Url) Then Call Execute(sAction_Plugin_TCategory_Url)
			If bAction_Plugin_TCategory_Url=True Then Exit Property
		Next

		Url = ZC_BLOG_HOST & "catalog.asp?"& "cate=" & ID

		'If CheckPluginState("STACentre") Then
		'	'Dim objCate : Set objCate=New STACentre_Categorys : If objCate.LoadInfoByID(ID) Then Url=objCate.Url : Set objCate=Nothing
		'End If

		Call Filter_Plugin_TCategory_Url(Url)

	End Property

	Public Property Get RssUrl
		RssUrl = ZC_BLOG_HOST & "feed.asp?cate=" & ID
	End Property

	Public Property Get HtmlName
		HtmlName=TransferHTML(Name,"[html-format]")
	End Property

	Public Property Get HtmlUrl
		HtmlUrl=TransferHTML(Url,"[html-format]")
	End Property
	

	Public Function Post()

		Call Filter_Plugin_TCategory_Post(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,MetaString)

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(Order,"int",0)
		Call CheckParameter(ParentID,"int",0)

		'ID可以为0
		Name=FilterSQL(Name)
		Alias=TransferHTML(Alias,"[filename]")
		Alias=FilterSQL(Alias)
		Intro=FilterSQL(Intro)

		TemplateName=UCase(FilterSQL(TemplateName))
		If TemplateName="CATALOG" Then TemplateName=""

		If Len(Name)=0 Then Post=False:Exit Function

		If ID=0 Then

			If Not ParentID=0 Then
				Set objRS=objConn.Execute("SELECT [cate_ID],[cate_ParentID] FROM [blog_Category] WHERE [cate_ID]=" & ParentID)
				If (Not objRS.bof) And (Not objRS.eof) Then
					If Not objRS("cate_ParentID")=0 Then ShowError(51)
				Else
					ShowError(50)
				End If
			End If

			objConn.Execute("INSERT INTO [blog_Category]([cate_Name],[cate_Order],[cate_Intro],[cate_ParentID],[cate_Url],[cate_Template],[cate_Meta]) VALUES ('"&Name&"',"&Order&",'"&Intro&"',"&ParentID&",'"&Alias&"','"&TemplateName&"','"&MetaString&"')")

			Dim objRS
			Set objRS=objConn.Execute("SELECT MAX([cate_ID]) FROM [blog_Category]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
			End If
			Set objRS=Nothing

			If ParentID=ID Then
				ParentID=0
				objConn.Execute("UPDATE [blog_Category] set [cate_Name]='"&Name&"',[cate_Order]="&Order&",[cate_Intro]='"&Intro&"',[cate_ParentID]="&ParentID&",[cate_Url]='"&Alias&"',[cate_Template]='"&TemplateName&"',[cate_Meta]='"&MetaString&"' WHERE [cate_ID] =" & ID)
			End If

		Else

			'如果选择了父分类, 需要进行以下判断
			If Not ParentID=0 Then
				'父分类不能为自身
				If ParentID=ID Then ParentID=0
				'不能将分类置于子分类下, 兼判断选择的父分类是否存在.
				Set objRS=objConn.Execute("SELECT [cate_ID],[cate_ParentID] FROM [blog_Category] WHERE [cate_ID]=" & ParentID)
				If (Not objRS.bof) And (Not objRS.eof) Then
					If Not objRS("cate_ParentID")=0 Then ShowError(51)
				Else
					ShowError(50)
				End If
				'不能将含有子分类的分类置于其它分类下.
				Set objRS=objConn.Execute("SELECT [cate_ID] FROM [blog_Category] WHERE [cate_ParentID]=" & ID)
				If (Not objRS.bof) And (Not objRS.eof) Then  ShowError(51)
				Set objRS=Nothing
			End If

			objConn.Execute("UPDATE [blog_Category] set [cate_Name]='"&Name&"',[cate_Order]="&Order&",[cate_Intro]='"&Intro&"',[cate_ParentID]="&ParentID&",[cate_Url]='"&Alias&"',[cate_Template]='"&TemplateName&"',[cate_Meta]='"&MetaString&"' WHERE [cate_ID] =" & ID)

		End If

		'修复父分类为自身的情况
		'objConn.Execute("UPDATE [blog_Category] set [cate_ParentID]=0,[cate_Meta]='"&MetaString&"' WHERE [cate_ID] = [cate_ParentID]")

		Post=True

	End Function


	Public Function LoadInfoByID(cate_ID)

		Call CheckParameter(cate_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [cate_ID],[cate_Name],[cate_Intro],[cate_Order],[cate_Count],[cate_ParentID],[cate_Url],[cate_Template],[cate_Meta] FROM [blog_Category] WHERE [cate_ID]=" & cate_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("cate_ID")
			Name=objRS("cate_Name")
			Alias=objRS("cate_Url")
			Order=objRS("cate_Order")
			Count=objRS("cate_Count")
			ParentID=objRS("cate_ParentID")
			Intro=objRS("cate_Intro")
			TemplateName=objRS("cate_Template")
			MetaString=objRS("cate_Meta")
			LoadInfoByID=True

		End If

		objRS.Close
		Set objRS=Nothing

		Call Filter_Plugin_TCategory_LoadInfoByID(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,MetaString)

	End Function


	Public Function LoadInfoByArray(aryCateInfo)

		If IsArray(aryCateInfo)=True Then
			ID=aryCateInfo(0)
			Name=aryCateInfo(1)
			Intro=aryCateInfo(2)
			Order=aryCateInfo(3)
			Count=aryCateInfo(4)
			ParentID=aryCateInfo(5)
			Alias=aryCateInfo(6)
			TemplateName=aryCateInfo(7)
			MetaString=aryCateInfo(8)
		End If

		LoadInfoByArray=True

		Call Filter_Plugin_TCategory_LoadInfoByArray(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,MetaString)

	End Function


	Public Function Del()

		Call Filter_Plugin_TCategory_Del(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,MetaString)

		Call CheckParameter(ID,"int",0)
		If (ID=0) Then Del=False:Exit Function
		Dim objRS
		Set objRS=objConn.Execute("SELECT [log_ID] FROM [blog_Article] WHERE [log_CateID]=" & ID)
		If (Not objRS.bof) And (Not objRS.eof) Then  ShowError(13)

		'不能删除含有子分类的分类.
		Set objRS=objConn.Execute("SELECT [cate_ID] FROM [blog_Category] WHERE [cate_ParentID]=" & ID)
		If (Not objRS.bof) And (Not objRS.eof) Then  ShowError(49)

		objConn.Execute("DELETE FROM [blog_Category] WHERE [cate_ID] =" & ID)

		Set objRS=Nothing

		Del=True

	End Function

	Private Sub Class_Initialize()
		ID=0
		Set Meta=New TMeta
	End Sub

End Class
'*********************************************************




'*********************************************************
' 目的：    定义TArticle类
' 输入：    无
' 返回：    无
'*********************************************************
Class TArticle

	Public ID
	Public CateID
	Public AuthorID
	Public Level

	Public Title
	Public Intro
	Public Content
	Public PostTime

	Public Tag
	Public Alias

	Public CommNums
	Public ViewNums
	Public TrackBackNums

	Private IP

	Public Istop
	Public FullUrl
	Public Meta
	Public TemplateName

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property

	Public Template_Article_Trackback
	Public Template_Article_Comment
	Public Template_Article_Commentpost
	Public Template_Article_Tag
	Public Template_Article_Navbar_L
	Public Template_Article_Navbar_R
	Public Template_Article_Commentpost_Verify
	Public Template_Article_Mutuality


	Public Template_Article_Single
	Public Template_Article_Multi
	Public Template_Article_Istop
	Public Template_Article_Search

	Public html
	Public htmlWAP
	Public Template_Article_Multi_WAP
	Public Template_Article_Single_WAP

	Private Ftemplate_Wap
	Private Ftemplate

	Public Property Let template(strFileName)
		Ftemplate=GetTemplate("TEMPLATE_" & strFileName)
	End Property
	Public Property Get template
		If TemplateName<>"" Then
			Dim s
			s=GetTemplate("TEMPLATE_" &TemplateName)
			If s<>"" Then
				template = s
			Else
				template = Ftemplate
			End If
		Else
			template = Ftemplate
		End If
	End Property

	Public Property Let template_Wap(strFileName)
		Ftemplate=GetTemplate("TEMPLATE_" & strFileName)
	End Property
	Public Property Get template_Wap
		template_Wap = Ftemplate_Wap
	End Property

	Private FDirectory
	Public Property Let Directory(strDirectory)
		FDirectory=strDirectory
	End Property
	Public Property Get Directory
		If IsEmpty(FDirectory)=True Then
			If ZC_CUSTOM_DIRECTORY_ENABLE=True Then
				Directory=ParseCustomDirectory(ZC_CUSTOM_DIRECTORY_REGEX,ZC_STATIC_DIRECTORY,Categorys(CateID).StaticName,Users(AuthorID).StaticName,Year(PostTime),Month(PostTime),Day(PostTime),ID,StaticName)
			Else
				Directory=ZC_STATIC_DIRECTORY
			End If
		Else
			Directory = FDirectory
		End If
		Directory=Replace(Directory,"\","/")
		If Right(ZC_BLOG_HOST & Directory,1)<>"/" Then
			Directory=Directory & "/"
		End If
	End Property


	Public Property Get Url

		'plugin node
		bAction_Plugin_TArticle_Url=False
		For Each sAction_Plugin_TArticle_Url in Action_Plugin_TArticle_Url
			If Not IsEmpty(sAction_Plugin_TArticle_Url) Then Call Execute(sAction_Plugin_TArticle_Url)
			If bAction_Plugin_TArticle_Url=True Then Exit Property
		Next

		If Level<=2 Then
			Url = ZC_BLOG_HOST & "view.asp?id=" & ID
		Else
			Url = ZC_BLOG_HOST & Directory & FileName
			If ZC_CUSTOM_DIRECTORY_ENABLE And ZC_CUSTOM_DIRECTORY_ANONYMOUS Then
				Url = ZC_BLOG_HOST & Directory
			End If
		End If

		Call Filter_Plugin_TArticle_Url(Url)

	End Property

	Public Property Get StaticName
		If IsNull(Alias) Or IsEmpty(Alias) Or Alias="" Then
			StaticName = ID
		Else
			StaticName = Alias
		End If
	End Property

	Public Property Get FileName

		FileName = StaticName

		If ZC_CUSTOM_DIRECTORY_ENABLE And ZC_CUSTOM_DIRECTORY_ANONYMOUS Then
			FileName = "default"
		End If
		FileName = FileName & "." & ZC_STATIC_TYPE
	End Property

	Private FTrackBackKey
	Public Property Get TrackBackKey
		If IsNull(FTrackBackKey) Or IsEmpty(FTrackBackKey) Or FTrackBackKey="" Then
			FTrackBackKey=Left(MD5(ZC_BLOG_HOST & ZC_BLOG_CLSID & CStr(ID) & CStr(TrackBackNums)),8)
		End If
		TrackBackKey=FTrackBackKey
	End Property

	Private FCommentKey
	Public Property Get CommentKey
		If IsNull(FCommentKey) Or IsEmpty(FCommentKey) Or FCommentKey="" Then
			FCommentKey=Left(MD5(ZC_BLOG_HOST & ZC_BLOG_CLSID & CStr(ID)),8)
		End If
		CommentKey=FCommentKey
	End Property

	Public Property Get TrackBack
		TrackBack = ZC_BLOG_HOST & "zb_system/cmd.asp?act=tb&id="& ID &"&key=" & TrackBackKey
	End Property

	Public Property Get PreTrackBack
		PreTrackBack = ZC_BLOG_HOST & "zb_system/cmd.asp?act=gettburl&id=" & ID
	End Property

	Public Property Get TrackBackUrl
		TrackBackUrl = TrackBack
	End Property

	Public Property Get CommentUrl
		CommentUrl = Url & "#comment"
	End Property

	Public Property Get WfwComment
		WfwComment = ZC_BLOG_HOST
	End Property

	Public Property Get WfwCommentRss
		WfwCommentRss = ZC_BLOG_HOST & "feed.asp?cmt=" & ID
	End Property

	Public Property Get WAPUrl
		WAPUrl = ZC_BLOG_HOST & "wap.asp?act=View&id=" & ID
	End Property

	Public Property Get HtmlWAPUrl
		HtmlWAPUrl=TransferHTML(WAPUrl,"[html-format]")
	End Property

	Public Property Get CommentPostUrl
		CommentPostUrl = ZC_BLOG_HOST & "zb_system/cmd.asp?act=cmt&key=" & CommentKey
	End Property

	Public Property Get HtmlContent
		HtmlContent=TransferHTML(UBBCode(Content,"[face][link][email][autolink][font][code][image][typeset][media][flash][key]"),"[html-japan][upload]")
	End Property

	Public Property Get HtmlTitle
		HtmlTitle=TransferHTML(Title,"[html-japan][html-format]")
	End Property

	Public Property Get HtmlIntro
		HtmlIntro=TransferHTML(UBBCode(Intro,"[face][link][email][autolink][font][code][image][typeset][media][flash][key]"),"[html-japan][upload]")
	End Property

	Public Property Get HtmlUrl
		HtmlUrl=TransferHTML(Url,"[html-format]")
	End Property

	Public Property Get TagToName

		Dim t,i,s

		If Tag<>"" Then
			s=Tag
			s=Replace(s,"}","")
			t=Split(s,"{")

			For i=LBound(t) To UBound(t)
				If t(i)<>"" Then
					If IsEmpty(FirstTagIntro) Then FirstTagIntro=Tags(t(i)).Intro
					t(i)=Tags(t(i)).Name
				End If
			Next

			s=Join(t,",")
			s=Right(s,Len(s)-1)

			TagToName=s
		End If

	End Property

	Public FirstTagIntro

	Public Function LoadInfobyID(log_ID)

		Call CheckParameter(log_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Meta] FROM [blog_Article] WHERE [log_ID]=" & log_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("log_ID")
			Tag=objRS("log_Tag")
			CateID=objRS("log_CateID")
			Title=objRS("log_Title")
			Intro=objRS("log_Intro")
			Content=objRS("log_Content")
			Level=objRS("log_Level")
			AuthorID=objRS("log_AuthorID")
			PostTime=objRS("log_PostTime")
			CommNums=objRS("log_CommNums")
			ViewNums=objRS("log_ViewNums")
			TrackBackNums=objRS("log_TrackBackNums")
			Alias=objRS("log_Url")
			Istop=objRS("log_Istop")
			TemplateName=objRS("log_Template")
			FullUrl=objRS("log_FullUrl")
			MetaString=objRS("log_Meta")

			Content=TransferHTML(Content,"[upload][zc_blog_host]")
			Intro=TransferHTML(Intro,"[upload][zc_blog_host]")

			PostTime = Year(PostTime) & "-" & Month(PostTime) & "-" & Day(PostTime) & " " & Hour(PostTime) & ":" & Minute(PostTime) & ":" & Second(PostTime)

		Else
			Exit Function
		End If

		objRS.Close
		Set objRS=Nothing

		LoadInfobyID=True

		Call Filter_Plugin_TArticle_LoadInfobyID(ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,FullUrl,MetaString)

	End Function


	Public Function LoadInfoByArray(aryArticleInfo)

		'[log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url]
		'Array(objRS("log_ID"),objRS("log_Tag"),objRS("log_CateID"),objRS("log_Title"),objRS("log_Intro"),objRS("log_Content"),objRS("log_Level"),objRS("log_AuthorID"),objRS("log_PostTime"),objRS("log_CommNums"),objRS("log_ViewNums"),objRS("log_TrackBackNums"),objRS("log_Url"))

		If IsArray(aryArticleInfo)=True Then

			ID=aryArticleInfo(0)
			Tag=aryArticleInfo(1)
			CateID=aryArticleInfo(2)
			Title=aryArticleInfo(3)
			Intro=aryArticleInfo(4)
			Content=aryArticleInfo(5)
			Level=aryArticleInfo(6)
			AuthorID=aryArticleInfo(7)
			PostTime=aryArticleInfo(8)
			CommNums=aryArticleInfo(9)
			ViewNums=aryArticleInfo(10)
			TrackBackNums=aryArticleInfo(11)
			Alias=aryArticleInfo(12)
			Istop=aryArticleInfo(13)
			TemplateName=aryArticleInfo(14)
			FullUrl=aryArticleInfo(15)
			MetaString=aryArticleInfo(16)

			Content=TransferHTML(Content,"[upload][zc_blog_host]")
			Intro=TransferHTML(Intro,"[upload][zc_blog_host]")

			PostTime = Year(PostTime) & "-" & Month(PostTime) & "-" & Day(PostTime) & " " & Hour(PostTime) & ":" & Minute(PostTime) & ":" & Second(PostTime)

		End If

		LoadInfoByArray=True

		Call Filter_Plugin_TArticle_LoadInfoByArray(ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,FullUrl,MetaString)

	End Function


	Public Function Export(intType)

		'plugin node
		bAction_Plugin_TArticle_Export_Begin=False
		For Each sAction_Plugin_TArticle_Export_Begin in Action_Plugin_TArticle_Export_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_Begin)
			If bAction_Plugin_TArticle_Export_Begin=True Then Exit Function
		Next

		Call Export_Tag
		Call Export_CMTandTB
		Call Export_CommentPost
		Call Export_Mutuality
		Call Export_NavBar

		Template_Article_Single=GetTemplate("TEMPLATE_B_ARTICLE-SINGLE")
		Template_Article_Multi=GetTemplate("TEMPLATE_B_ARTICLE-MULTI")
		Template_Article_Istop=GetTemplate("TEMPLATE_B_ARTICLE-ISTOP")
		Template_Article_Multi_WAP=GetTemplate("TEMPLATE_WAP_ARTICLE-MULTI")
		Template_Article_Single_WAP =GetTemplate("TEMPLATE_WAP_SINGLE")

		'plugin node
		Call Filter_Plugin_TArticle_Export_Template(Ftemplate,Template_Article_Single,Template_Article_Multi,Template_Article_Istop,Template_Article_Multi_WAP,Template_Article_Single_WAP)
		'plugin node
		Call Filter_Plugin_TArticle_Export_Template_Sub(Template_Article_Comment,Template_Article_Trackback,Template_Article_Tag,Template_Article_Commentpost,Template_Article_Navbar_L,Template_Article_Navbar_R,Template_Article_Mutuality)

		Template_Article_Single=Replace(Template_Article_Single,"<#template:article_trackback#>",Template_Article_Trackback)
		Template_Article_Single=Replace(Template_Article_Single,"<#template:article_comment#>",Template_Article_Comment)
		Template_Article_Single=Replace(Template_Article_Single,"<#template:article_commentpost#>",Template_Article_Commentpost)
		Template_Article_Single=Replace(Template_Article_Single,"<#template:article_tag#>",Template_Article_Tag)
		Template_Article_Single=Replace(Template_Article_Single,"<#template:article_navbar_l#>",Template_Article_Navbar_L)
		Template_Article_Single=Replace(Template_Article_Single,"<#template:article_navbar_r#>",Template_Article_Navbar_R)
		Template_Article_Single=Replace(Template_Article_Single,"<#template:article_mutuality#>",Template_Article_Mutuality)

		Template_Article_Multi=Replace(Template_Article_Multi,"<#template:article_tag#>",Template_Article_Tag)
		Template_Article_Istop=Replace(Template_Article_Istop,"<#template:article_tag#>",Template_Article_Tag)

		Ftemplate=Replace(Ftemplate,"<#template:article_trackback#>",Template_Article_Trackback)
		Ftemplate=Replace(Ftemplate,"<#template:article_comment#>",Template_Article_Comment)
		Ftemplate=Replace(Ftemplate,"<#template:article_commentpost#>",Template_Article_Commentpost)
		Ftemplate=Replace(Ftemplate,"<#template:article_tag#>",Template_Article_Tag)
		Ftemplate=Replace(Ftemplate,"<#template:article_navbar_l#>",Template_Article_Navbar_L)
		Ftemplate=Replace(Ftemplate,"<#template:article_navbar_r#>",Template_Article_Navbar_R)
		Ftemplate=Replace(Ftemplate,"<#template:article_mutuality#>",Template_Article_Mutuality)

		Dim aryTemplateTagsName()
		Dim aryTemplateTagsValue()
		Dim i,j
		ReDim aryTemplateTagsName(49)
		ReDim aryTemplateTagsValue(49)

		aryTemplateTagsName(1)="article/id"
		aryTemplateTagsValue(1)=ID
		aryTemplateTagsName(2)="article/level"
		aryTemplateTagsValue(2)=Level
		aryTemplateTagsName(3)="article/title"
		If intType=ZC_DISPLAY_MODE_SEARCH Then
			aryTemplateTagsValue(3)=Search(Title,Request.QueryString("q"))
		Else
			aryTemplateTagsValue(3)=HtmlTitle
		End If
		aryTemplateTagsName(4)="article/intro"
		If intType=ZC_DISPLAY_MODE_SEARCH Then
			'aryTemplateTagsValue(4)=Search(TransferHTML(Intro & Content,"[html-format]"),Request.QueryString("q"))
			aryTemplateTagsValue(4)=Search(TransferHTML(Intro & Content,"[nohtml]"),Request.QueryString("q"))
		Else
			If Level=2 Then
				aryTemplateTagsValue(4)=ZC_MSG043
			Else
				aryTemplateTagsValue(4)=HtmlIntro
			End If
		End If
		aryTemplateTagsName(5)="article/content"
		aryTemplateTagsValue(5)=HtmlContent
		If intType=ZC_DISPLAY_MODE_SEARCH Then
			aryTemplateTagsValue(5)=aryTemplateTagsValue(4)
		End If
		aryTemplateTagsName(6)="article/posttime"
		aryTemplateTagsValue(6)=PostTime
		aryTemplateTagsName(7)="article/commnums"
		aryTemplateTagsValue(7)=Commnums
		aryTemplateTagsName(8)="article/viewnums"
		aryTemplateTagsValue(8)=Viewnums
		aryTemplateTagsName(9)="article/trackbacknums"
		aryTemplateTagsValue(9)=Trackbacknums
		aryTemplateTagsName(10)="article/trackback_url"
		aryTemplateTagsValue(10)=TrackBack
		aryTemplateTagsName(11)="article/url"
		aryTemplateTagsValue(11)=HtmlUrl

		aryTemplateTagsName(12)="article/category/id"
		aryTemplateTagsValue(12)=Categorys(CateID).ID
		aryTemplateTagsName(13)="article/category/name"
		aryTemplateTagsValue(13)=Categorys(CateID).HtmlName
		aryTemplateTagsName(15)="article/category/order"
		aryTemplateTagsValue(15)=Categorys(CateID).Order
		aryTemplateTagsName(16)="article/category/count"
		aryTemplateTagsValue(16)=Categorys(CateID).Count
		aryTemplateTagsName(17)="article/category/url"
		aryTemplateTagsValue(17)=Categorys(CateID).HtmlUrl

		aryTemplateTagsName(18)="article/author/id"
		aryTemplateTagsValue(18)=Users(AuthorID).ID
		aryTemplateTagsName(19)="article/author/name"
		aryTemplateTagsValue(19)=Users(AuthorID).Name
		aryTemplateTagsName(20)="article/author/level"
		aryTemplateTagsValue(20)=ZVA_User_Level_Name(Users(AuthorID).Level)
		aryTemplateTagsName(21)="article/author/email"
		aryTemplateTagsValue(21)=Users(AuthorID).Email
		aryTemplateTagsName(22)="article/author/homepage"
		aryTemplateTagsValue(22)=Users(AuthorID).HomePage
		aryTemplateTagsName(23)="article/author/count"
		aryTemplateTagsValue(23)=Users(AuthorID).Count
		aryTemplateTagsName(24)="article/author/url"
		aryTemplateTagsValue(24)=Users(AuthorID).HtmlUrl

		aryTemplateTagsName(25)="article/posttime/longdate"
		aryTemplateTagsValue(25)=FormatDateTime(PostTime,vbLongDate)
		aryTemplateTagsName(26)="article/posttime/shortdate"
		aryTemplateTagsValue(26)=FormatDateTime(PostTime,vbShortDate)
		aryTemplateTagsName(27)="article/posttime/longtime"
		aryTemplateTagsValue(27)=FormatDateTime(PostTime,vbLongTime)
		aryTemplateTagsName(28)="article/posttime/shorttime"
		aryTemplateTagsValue(28)=FormatDateTime(PostTime,vbShortTime)
		aryTemplateTagsName(29)="article/posttime/year"
		aryTemplateTagsValue(29)=Year(PostTime)
		aryTemplateTagsName(30)="article/posttime/month"
		aryTemplateTagsValue(30)=Month(PostTime)
		aryTemplateTagsName(31)="article/posttime/monthname"
		aryTemplateTagsValue(31)=ZVA_Month(Month(PostTime))
		aryTemplateTagsName(32)="article/posttime/day"
		aryTemplateTagsValue(32)=Day(PostTime)
		aryTemplateTagsName(33)="article/posttime/weekday"
		aryTemplateTagsValue(33)=Weekday(PostTime)
		aryTemplateTagsName(34)="article/posttime/weekdayname"
		aryTemplateTagsValue(34)=ZVA_Week(Weekday(PostTime))
		aryTemplateTagsName(35)="article/posttime/hour"
		aryTemplateTagsValue(35)=Hour(PostTime)
		aryTemplateTagsName(36)="article/posttime/minute"
		aryTemplateTagsValue(36)=Minute(PostTime)
		aryTemplateTagsName(37)="article/posttime/second"
		aryTemplateTagsValue(37)=Second(PostTime)

		aryTemplateTagsName(38)="article/commentrss"
		aryTemplateTagsValue(38)=WfwCommentRss
		aryTemplateTagsName(39)="article/commentposturl"
		aryTemplateTagsValue(39)=TransferHTML(CommentPostUrl,"[html-format]")
		aryTemplateTagsName(40)="article/pretrackback_url"
		aryTemplateTagsValue(40)=TransferHTML(PreTrackBack,"[html-format]")
		aryTemplateTagsName(41)="article/trackbackkey"
		aryTemplateTagsValue(41)=TrackBackKey
		aryTemplateTagsName(42)="article/commentkey"
		aryTemplateTagsValue(42)=CommentKey

		aryTemplateTagsName(43)="article/staticname"
		aryTemplateTagsValue(43)=StaticName
		aryTemplateTagsName(44)="article/category/staticname"
		aryTemplateTagsValue(44)=""'Categorys(CateID).StaticName
		aryTemplateTagsName(45)="article/author/staticname"
		aryTemplateTagsValue(45)=""'Users(AuthorID).StaticName
		aryTemplateTagsName(46)="article/tagtoname"
		aryTemplateTagsValue(46)=TagToName
		aryTemplateTagsName(47)="article/firsttagintro"
		aryTemplateTagsValue(47)=FirstTagIntro

		aryTemplateTagsName(48)="article/posttime/monthnameabbr"
		aryTemplateTagsValue(48)=ZVA_Month_Abbr(Month(PostTime))
		aryTemplateTagsName(49)="article/posttime/weekdaynameabbr"
		aryTemplateTagsValue(49)=ZVA_Week_Abbr(Weekday(PostTime))


		Call Filter_Plugin_TArticle_Export_TemplateTags(aryTemplateTagsName,aryTemplateTagsValue)

		j=UBound(aryTemplateTagsName)
		For i=1 to j
			If IsNull(aryTemplateTagsValue(i))=False Then
				Template_Article_Istop=Replace(Template_Article_Istop,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
				Template_Article_Multi=Replace(Template_Article_Multi,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
				Template_Article_Single=Replace(Template_Article_Single,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
				Template_Article_Multi_WAP = Replace(Template_Article_Multi_WAP,"<#" & aryTemplateTagsName(i) & "#>", aryTemplateTagsValue(i))
				Template_Article_Single_WAP = Replace(Template_Article_Single_WAP,"<#" & aryTemplateTagsName(i) & "#>", aryTemplateTagsValue(i))
				Ftemplate = Replace(Ftemplate,"<#" & aryTemplateTagsName(i) & "#>", aryTemplateTagsValue(i))
			Else
				Template_Article_Istop=Replace(Template_Article_Istop,"<#" & aryTemplateTagsName(i) & "#>","")
				Template_Article_Multi=Replace(Template_Article_Multi,"<#" & aryTemplateTagsName(i) & "#>","")
				Template_Article_Single=Replace(Template_Article_Single,"<#" & aryTemplateTagsName(i) & "#>","")
				Template_Article_Multi_WAP = Replace(Template_Article_Multi_WAP,"<#" & aryTemplateTagsName(i) & "#>", "")
				Template_Article_Single_WAP = Replace(Template_Article_Single_WAP,"<#" & aryTemplateTagsName(i) & "#>", "")
				Ftemplate = Replace(Ftemplate,"<#" & aryTemplateTagsName(i) & "#>", "")
			End If
		Next

		If intType=ZC_DISPLAY_MODE_SEARCH Then
			Template_Article_Search=Template_Article_Multi
		End If

		Export=True

		'plugin node
		bAction_Plugin_TArticle_Export_End=False
		For Each sAction_Plugin_TArticle_Export_End in Action_Plugin_TArticle_Export_End
			If Not IsEmpty(sAction_Plugin_TArticle_Export_End) Then Call Execute(sAction_Plugin_TArticle_Export_End)
			If bAction_Plugin_TArticle_Export_End=True Then Exit Function
		Next

	End Function


	Public Function Export_Tag

		'plugin node
		bAction_Plugin_TArticle_Export_Tag_Begin=False
		For Each sAction_Plugin_TArticle_Export_Tag_Begin in Action_Plugin_TArticle_Export_Tag_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_Tag_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_Tag_Begin)
			If bAction_Plugin_TArticle_Export_Tag_Begin=True Then Exit Function
		Next

		'Tag
		Dim t,i,s,j

		If Tag<>"" Then

			Tag=Replace(Tag,"}","")
			t=Split(Tag,"{")

			For i=LBound(t) To UBound(t)
				If t(i)<>"" Then

					s=GetTemplate("TEMPLATE_B_ARTICLE_TAG")

					Template_Article_Tag=Template_Article_Tag & Tags(t(i)).MakeTemplate(s)

				End If
			Next

		End If

		Export_Tag=True

	End Function


	Function Export_CMTandTB()

		'plugin node
		bAction_Plugin_TArticle_Export_CMTandTB_Begin=False
		For Each sAction_Plugin_TArticle_Export_CMTandTB_Begin in Action_Plugin_TArticle_Export_CMTandTB_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_CMTandTB_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_CMTandTB_Begin)
			If bAction_Plugin_TArticle_Export_CMTandTB_Begin=True Then Exit Function
		Next

		If CommNums + TrackBackNums > 0 Then
			Dim strC_Count,strC,strT_Count,strT

			Dim objComment
			Dim objTrackBack

			Dim i

			Dim objRS
			strC_Count=0
			Set objRS=Server.CreateObject("ADODB.Recordset")
			objRS.CursorType = adOpenKeyset
			objRS.LockType = adLockReadOnly
			objRS.ActiveConnection=objConn
			objRS.Source="SELECT [comm_ID],[log_ID],[comm_AuthorID],[comm_Author],[comm_Content],[comm_Email],[comm_HomePage],[comm_PostTime],[comm_IP],[comm_Agent],[comm_Reply],[comm_LastReplyIP],[comm_LastReplyTime],[comm_ParentID],[comm_ParentCount],[comm_Meta] FROM [blog_Comment] WHERE ([blog_Comment].[log_ID]=" & ID &" AND [blog_Comment].[comm_ParentID]=0) UNION ALL SELECT [tb_ID],[log_ID],'',[tb_Title],[tb_Excerpt],[tb_Blog],[tb_URL],[tb_PostTime],[tb_IP],[tb_Agent],'','','','' ,'',[tb_Meta] from [blog_TrackBack] WHERE [blog_TrackBack].[log_ID]="& ID & " ORDER BY [comm_ID],[comm_PostTime]"
			objRS.Open()

			If (not objRS.bof) And (not objRS.eof) Then

				ReDim aryArticleExportMsgTB(objRS.RecordCount)

				For i=1 To objRS.RecordCount

					If IsNumeric(objRS("comm_AuthorID")) Then

						Set objComment=New TComment
						objComment.LoadInfoByArray(Array(objRS("comm_ID"),objRS("log_ID"),objRS("comm_AuthorID"),objRS("comm_Author"),objRS("comm_Content"),objRS("comm_Email"),objRS("comm_HomePage"),objRS("comm_PostTime"),"","",objRS("comm_Reply"),"","",objRS("comm_ParentID"),objRS("comm_ParentCount"),objRS("comm_Meta")))

						If objRs("comm_ParentID")=0 Then strC_Count=strC_Count+1

						strC=GetTemplate("TEMPLATE_B_ARTICLE_COMMENT")
'Mark3
						objComment.Count=strC_Count
						
						strC=objComment.MakeTemplate(strC,False)

						If ZC_COMMENT_REVERSE_ORDER_EXPORT=True Then
							Template_Article_Comment=strC & Template_Article_Comment
						Else
							Template_Article_Comment=Template_Article_Comment & strC
						End If

						Set objComment=Nothing

					Else

						Set objTrackBack=New TTrackBack

						objTrackBack.LoadInfoByArray(Array(objRS("comm_ID"),objRS("log_ID"),objRS("comm_HomePage"),objRS("comm_Author"),objRS("comm_Email"),objRS("comm_Content"),objRS("comm_PostTime"),"","","",objRS("comm_Meta")))

						strT_Count=strT_Count+1

						strT=GetTemplate("TEMPLATE_B_ARTICLE_TRACKBACK")

						objTrackBack.Count=strT_Count
						strT=objTrackBack.MakeTemplate(strT)

						If ZC_COMMENT_REVERSE_ORDER_EXPORT=True Then
							Template_Article_Trackback=strT & Template_Article_Trackback
						Else
							Template_Article_Trackback=Template_Article_Trackback & strT
						End If

						Set objTrackBack=Nothing

					End If

					objRS.MoveNext
					If objRS.eof Then Exit For
				Next

			End if

			objRS.Close()
			Set objRS=Nothing

		End If

		Template_Article_Comment=Template_Article_Comment & "<div style=""display:none;"" id=""divAjaxComment""></div>"

		Export_CMTandTB=True

	End Function


	Function Export_NavBar()

		'plugin node
		bAction_Plugin_TArticle_Export_NavBar_Begin=False
		For Each sAction_Plugin_TArticle_Export_NavBar_Begin in Action_Plugin_TArticle_Export_NavBar_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_NavBar_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_NavBar_Begin)
			If bAction_Plugin_TArticle_Export_NavBar_Begin=True Then Exit Function
		Next


		If ZC_USE_NAVIGATE_ARTICLE=False Or CateID=0 Then

			Template_Article_Navbar_L=""
			Template_Article_Navbar_R=""

			Export_NavBar=True
			Exit Function

		End If

		Dim s,t

		Dim strName
		Dim strUrl

		Dim objNavArticle

		Dim objRS

		Set objRS=objConn.Execute("SELECT TOP 1 [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Meta] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_Level]>2) AND ([log_PostTime]<" & ZC_SQL_POUND_KEY & PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] DESC")
		If (Not objRS.bof) And (Not objRS.eof) Then

			Set objNavArticle=New TArticle
			If objNavArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16))) Then
				strName=objNavArticle.Title
				strUrl=objNavArticle.Url
			End If
			Set objNavArticle=Nothing
			s=GetTemplate("TEMPLATE_B_ARTICLE_NVABAR_L")
			
			s=Replace(s,"<#article/nav_l/url#>",strUrl)
			s=Replace(s,"<#article/nav_l/name#>",strName)

			Template_Article_Navbar_L=s

		End If
		Set objRS=Nothing

		Set objRS=objConn.Execute("SELECT TOP 1 [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Meta] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_Level]>2) AND ([log_PostTime]>" & ZC_SQL_POUND_KEY & PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] ASC")
		If (Not objRS.bof) And (Not objRS.eof) Then

			Set objNavArticle=New TArticle
			If objNavArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16))) Then
				strName=objNavArticle.Title
				strUrl=objNavArticle.Url
			End If
			Set objNavArticle=Nothing

			t=GetTemplate("TEMPLATE_B_ARTICLE_NVABAR_R")

			t=Replace(t,"<#article/nav_r/url#>",strUrl)
			t=Replace(t,"<#article/nav_r/name#>",strName)

			Template_Article_Navbar_R=t

		End If
		Set objRS=Nothing

		Export_NavBar=True

	End Function


	Function Export_CommentPost()

		'plugin node
		bAction_Plugin_TArticle_Export_CommentPost_Begin=False
		For Each sAction_Plugin_TArticle_Export_CommentPost_Begin in Action_Plugin_TArticle_Export_CommentPost_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_CommentPost_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_CommentPost_Begin)
			If bAction_Plugin_TArticle_Export_CommentPost_Begin=True Then Exit Function
		Next

		If Level<4 Then Exit Function

		Template_Article_Commentpost=GetTemplate("TEMPLATE_B_ARTICLE_COMMENTPOST")

		If ZC_COMMENT_VERIFY_ENABLE=True Then
			Template_Article_Commentpost_Verify=GetTemplate("TEMPLATE_B_ARTICLE_COMMENTPOST-VERIFY")
		End If

		Template_Article_Commentpost=Replace(Template_Article_Commentpost,"<#template:article_commentpost-verify#>",Template_Article_Commentpost_Verify)

	End Function


	'相关文章的生成
	Function Export_Mutuality()

		'plugin node
		bAction_Plugin_TArticle_Export_Mutuality_Begin=False
		For Each sAction_Plugin_TArticle_Export_Mutuality_Begin in Action_Plugin_TArticle_Export_Mutuality_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_Mutuality_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_Mutuality_Begin)
			If bAction_Plugin_TArticle_Export_Mutuality_Begin=True Then Exit Function
		Next

		If Tag<>"" Then

			Dim strCC_Count,strCC_ID,strCC_Name,strCC_Url,strCC_PostTime,strCC_Title
			Dim strCC
			Dim i
			Dim j
			Dim objRS
			Dim strSQL

			Set objRS=Server.CreateObject("ADODB.Recordset")

			strSQL="SELECT TOP "& ZC_MUTUALITY_COUNT &" [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Meta] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_Level]>2) AND [log_ID]<"& ID
			strSQL = strSQL & " AND ("

			Dim aryTAGs
			Tag=Replace(Tag,"}","")
			aryTAGs=Split(Tag,"{")

			For j = LBound(aryTAGs) To UBound(aryTAGs)
				If aryTAGs(j)<>"" Then
					strSQL = strSQL & "([log_Tag] Like '%{"&FilterSQL(aryTAGs(j))&"}%')"
					If j=UBound(aryTAGs) Then Exit For
					If aryTAGs(j)<>"" Then strSQL = strSQL & " OR "
				End If
			Next

			strSQL = strSQL & ")"
			strSQL = strSQL + " ORDER BY [log_PostTime] DESC "

			Set objRS=Server.CreateObject("ADODB.Recordset")
			objRS.CursorType = adOpenKeyset
			objRS.LockType = adLockReadOnly
			objRS.ActiveConnection=objConn
			objRS.Source=strSQL
			objRS.Open()
			If (Not objRS.bof) And (Not objRS.eof) Then

				Dim objArticle
				For i=1 To ZC_MUTUALITY_COUNT '相关文章数目，可自行设定

					Set objArticle=New TArticle

					If objArticle.LoadInfoByArray(Array(objRS("log_ID"),objRS("log_Tag"),objRS("log_CateID"),objRS("log_Title"),"","",objRS("log_Level"),objRS("log_AuthorID"),objRS("log_PostTime"),objRS("log_CommNums"),objRS("log_ViewNums"),objRS("log_TrackBackNums"),objRS("log_Url"),objRS("log_Istop"),objRS("log_Template"),objRS("log_FullUrl"),objRS("log_Meta"))) Then

						strCC_Count=strCC_Count+1
						strCC_ID=objArticle.ID
						strCC_Url=objArticle.Url
						strCC_PostTime=objArticle.PostTime
						strCC_Title=objArticle.Title

						strCC=GetTemplate("TEMPLATE_B_ARTICLE_Mutuality")

						strCC=Replace(strCC,"<#article/mutuality/id#>",strCC_ID)
						strCC=Replace(strCC,"<#article/mutuality/url#>",strCC_Url)
						strCC=Replace(strCC,"<#article/mutuality/posttime#>",strCC_PostTime)
						strCC=Replace(strCC,"<#article/mutuality/name#>",strCC_Title)

						Template_Article_Mutuality=Template_Article_Mutuality & strCC

					End If

					objRS.MoveNext
					If objRS.eof Then Exit For
					Set objArticle=Nothing
				Next

			End if

			objRS.Close()
			Set objRS=Nothing

		End If

		Export_Mutuality=True

	End Function



	Public Function Post()

		Call Filter_Plugin_TArticle_Post(ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,MetaString)

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(CateID,"int",0)
		Call CheckParameter(AuthorID,"int",0)
		Call CheckParameter(Level,"int",0)
		Call CheckParameter(PostTime,"dtm",Empty)
		Call CheckParameter(Istop,"bool",False)

		'ID可以为0
		'If (CateID=0) Then Post=False:Exit Function
		If (AuthorID=0) Then Post=False:Exit Function
		If IsEmpty(PostTime) Then Post=False:Exit Function

		Title=FilterSQL(Title)
		Intro=FilterSQL(Intro)
		Content=FilterSQL(Content)
		Tag=FilterSQL(Tag)
		IP=FilterSQL(IP)

		Title=TransferHTML(Title,"[japan-html]")
		Intro=TransferHTML(Intro,"[japan-html]")
		Content=TransferHTML(Content,"[japan-html]")

		Intro=TransferHTML(Intro,"[anti-upload]")
		Content=TransferHTML(Content,"[anti-upload]")

		'先进行"[anti-upload]"，再替换<#ZC_BLOG_HOST#>
		Intro=TransferHTML(Intro,"[anti-zc_blog_host]")
		Content=TransferHTML(Content,"[anti-zc_blog_host]")


		Alias=TransferHTML(Alias,"[filename]")
		Alias=FilterSQL(Alias)

		'检查“别名”是否有重名
		If Alias<>"" Then
			Dim objRSsub
			Set objRSsub=objConn.Execute("SELECT [log_ID] FROM [blog_Article] WHERE [log_ID]<>"& ID &" AND [log_Url]='"& Alias &"'" )
			If (Not objRSsub.bof) And (Not objRSsub.eof) Then
				Randomize
				Alias=Alias & "_" & CStr(Int((9 * Rnd) + 1)) & CStr(Int((9 * Rnd) + 1)) & CStr(Int((9 * Rnd) + 1)) & CStr(Int((9 * Rnd) + 1))
			End If
			Set objRSsub=Nothing
		End If

		If Len(Title)=0 Then Post=False:Exit Function
		If Len(Content)=0 Then Post=False:Exit Function
		If Len(Intro)=0 Then Intro=Left(Content,ZC_TB_EXCERPT_MAX) & "..."

		If Len(FullUrl)>255 Then FullUrl=Left(FullUrl,255)

		TemplateName=UCase(FilterSQL(TemplateName))
		If TemplateName="SINGLE" Then TemplateName=""
		If ID=0 Then
			objConn.Execute("INSERT INTO [blog_Article]([log_CateID],[log_AuthorID],[log_Level],[log_Title],[log_Intro],[log_Content],[log_PostTime],[log_IP],[log_Tag],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_ViewNums],[log_Meta]) VALUES ("&CateID&","&AuthorID&","&Level&",'"&Title&"','"&Intro&"','"&Content&"','"&PostTime&"','"&IP&"','"&Tag&"','"&Alias&"',"&CInt(Istop)&",'"&TemplateName&"','"&FullUrl&"',0,'"&MetaString&"')")
			Dim objRS
			Set objRS=objConn.Execute("SELECT MAX([log_ID]) FROM [blog_Article]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
			End If
			Set objRS=Nothing
		Else
			objConn.Execute("UPDATE [blog_Article] SET [log_CateID]="&CateID&",[log_AuthorID]="&AuthorID&",[log_Level]="&Level&",[log_Title]='"&Title&"',[log_Intro]='"&Intro&"',[log_Content]='"&Content&"',[log_PostTime]='"&PostTime&"',[log_IP]='"&IP&"',[log_Tag]='"&Tag&"',[log_Url]='"&Alias&"',[log_Istop]="&CInt(Istop)&",[log_Template]='"&TemplateName&"',[log_FullUrl]='"&FullUrl&"',[log_Meta]='"&MetaString&"' WHERE [log_ID] =" & ID)
		End If

		Post=True

	End Function


	Public Function DelFile()

		On Error Resume Next

		Dim fso, TxtFile

		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(BlogPath & Directory & FileName) Then
			Set TxtFile = fso.GetFile(BlogPath & Directory & FileName)
			TxtFile.Delete
		End If
		Set fso=Nothing

		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(BlogPath & "zb_users/cache/" & ID & ".html") Then
			Set TxtFile = fso.GetFile(BlogPath & "zb_users/cache/" & ID & ".html")
			TxtFile.Delete
		End If
		Set fso=Nothing

		DelFile=True

		Err.Clear

	End Function


	Public Function Del()

		Call Filter_Plugin_TArticle_Del(ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,MetaString)

		Call DelFile()

		Call CheckParameter(ID,"int",0)
		If (ID=0) Then Del=False:Exit Function

		objConn.Execute("DELETE FROM [blog_Article] WHERE [log_ID] =" & ID)
		objConn.Execute("DELETE FROM [blog_Comment] WHERE [log_ID] =" & ID)
		objConn.Execute("DELETE FROM [blog_TrackBack] WHERE [log_ID] =" & ID)

		Del=True

	End Function


	Public Function Statistic()

		Dim objRS
		Set objRS=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Comment] WHERE [log_ID] =" & ID)
		If (Not objRS.bof) And (Not objRS.eof) Then
			CommNums=objRS(0)
		End If
		objConn.Execute("UPDATE [blog_Article] SET [log_CommNums]="& CommNums &" WHERE [log_ID] =" & ID)
		Set objRS=Nothing

		Set objRS=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_TrackBack] WHERE [log_ID] =" & ID)
		If (Not objRS.bof) And (Not objRS.eof) Then
			TrackBackNums=objRS(0)
		End If
		objConn.Execute("UPDATE [blog_Article] SET [log_TrackBackNums]="& TrackBackNums &" WHERE [log_ID] =" & ID)
		Set objRS=Nothing

		Statistic=True

	End Function


	Function Build()

		Dim aryTemplateTagsName
		Dim aryTemplateTagsValue

		Dim i,j

		htmlWAP = Template_Article_Single_WAP

		If IsEmpty(html) Then html=template

		Call Filter_Plugin_TArticle_Build_Template(html,htmlWAP)
		Call Filter_Plugin_TArticle_Build_Template_Sub(Template_Article_Single)

		html=Replace(html,"<#template:article-single#>",Template_Article_Single)

		aryTemplateTagsName=TemplateTagsName
		aryTemplateTagsValue=TemplateTagsValue

		aryTemplateTagsName(0)="BlogTitle"
		aryTemplateTagsValue(0)=HtmlTitle

		Call Filter_Plugin_TArticle_Build_TemplateTags(aryTemplateTagsName,aryTemplateTagsValue)


		j=UBound(aryTemplateTagsName)

		For i=1 to j
			If (InStr(aryTemplateTagsName(i),"CACHE_INCLUDE_")>0) And (Right(aryTemplateTagsName(i),5)<>"_HTML") And (Right(aryTemplateTagsName(i),3)<>"_JS") Then

				Dim modname
				modname=LCase(Replace(aryTemplateTagsName(i),"CACHE_INCLUDE_",""))
				If aryTemplateTagsName(i)<>"CACHE_INCLUDE_NAVBAR" Then
					If aryTemplateTagsName(i)="CACHE_INCLUDE_CALENDAR" Then
						aryTemplateTagsValue(i)="<div id=""mod_"+modname+"""><script type=""text/javascript"">strBatchInculde+=""mod_"+modname+"="+modname+",""</script></div>"
					Else
						aryTemplateTagsValue(i)="<li id=""mod_"+modname+""" style=""display:none;""><script type=""text/javascript"">strBatchInculde+=""mod_"+modname+"="+modname+",""</script></li>"
					End If
				End If
			End If
		Next

		j=UBound(aryTemplateTagsName)

		For i=1 to j
			html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
			htmlWAP = Replace(htmlWAP, "<#" & aryTemplateTagsName(i) & "#>", aryTemplateTagsValue(i))
		Next
		html=Replace(html,"<#" & aryTemplateTagsName(0) & "#>",aryTemplateTagsValue(0))
		htmlWAP = Replace(htmlWAP, "<#" & aryTemplateTagsName(0) & "#>", aryTemplateTagsValue(0))

		Build=True

	End Function


	Function SetVar(TemplateTag,TemplateValue)

		If IsEmpty(html) Then html=template

		html=Replace(html,"<#" & TemplateTag & "#>",TemplateValue)

	End Function


	Function Save()

		If Not(Level>2) Then Save=True:Exit Function

		Dim objStream

		html=TransferHTML(html,"[no-asp]")
		If ZC_STATIC_TYPE="asp" Then
			html="<"&"%@ CODEPAGE=65001 %"&">" & html
		End If

		If ZC_CUSTOM_DIRECTORY_ENABLE=True Then
			Call CreatDirectoryByCustomDirectory(Directory)
		End If

		Call SaveToFile(BlogPath & Directory & FileName,html,"utf-8",False)

		Save=True

	End Function


	Function SaveCache()

		If Not(Level>1) Then SaveCache=True:Exit Function

		Dim strList

		If Istop Then
			strList=Template_Article_Istop
		Else
			strList=Template_Article_Multi
		End If
		strList=TransferHTML(strList,"[no-asp]")

		Call SaveToFile(BlogPath & "zb_users/cache/" & ID & ".html",strList,"utf-8",False)

		SaveCache=True

	End Function


	Function LoadCache()

		Dim objStream

		Template_Article_Multi=LoadFromFile(BlogPath & "zb_users/CACHE/" & ID & ".html","utf-8")

		LoadCache=True

	End Function


	Private Sub Class_Initialize()

		PostTime=GetTime(Now())
		ID=0
		CateID=0
		AuthorID=0
		Level=4'默认为普通
		Title=ZC_MSG099
		IP=Request.Servervariables("REMOTE_ADDR")

		Ftemplate_Wap=Empty
		Ftemplate=Empty

		Set Meta=New TMeta

	End Sub


End Class
'*********************************************************




'*********************************************************
' 目的：    定义TArticleList类
' 输入：    无
' 返回：    无
'*********************************************************
Class TArticleList

	Public Title

	Public FileName

	Public AllList
	Public AuthList
	Public CateList
	Public TagsList

	Public aryArticle
	Public aryArticleList()


	Public Template_PageBar
	Public Template_Article_Multi
	Public Template_PageBar_Next
	Public Template_PageBar_Previous
	Public Template_Calendar


	Public TemplateTags_ArticleList_Author_ID
	Public TemplateTags_ArticleList_Tags_ID
	Public TemplateTags_ArticleList_Category_ID
	Public TemplateTags_ArticleList_Date_ShortDate
	Public TemplateTags_ArticleList_Date_Year
	Public TemplateTags_ArticleList_Date_Month
	Public TemplateTags_ArticleList_Date_Day
	Public TemplateTags_ArticleList_Page_Now
	Public TemplateTags_ArticleList_Page_All

	Public html


	Private Ftemplate
	Public Property Let template(strFileName)
		Ftemplate=GetTemplate("TEMPLATE_" & strFileName)
	End Property
	Public Property Get template
		template = Ftemplate
	End Property


	Private FDirectory
	Public Property Let Directory(strDirectory)
		FDirectory=strDirectory
	End Property
	Public Property Get Directory
		If IsEmpty(FDirectory)=True Then
			Directory=ZC_STATIC_DIRECTORY
		Else
			Directory = FDirectory
		End If
		Directory=Replace(Directory,"\","/")
		If Right(ZC_BLOG_HOST & Directory,1)<>"/" Then
			Directory=Directory & "/"
		End If
	End Property


	Public Function Export(intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType)

		'plugin node
		bAction_Plugin_TArticleList_Export_Begin=False
		For Each sAction_Plugin_TArticleList_Export_Begin in Action_Plugin_TArticleList_Export_Begin
			If Not IsEmpty(sAction_Plugin_TArticleList_Export_Begin) Then Call Execute(sAction_Plugin_TArticleList_Export_Begin)
			If bAction_Plugin_TArticleList_Export_Begin=True Then Exit Function
		Next

		'plugin node
		Call Filter_Plugin_TArticleList_Export(intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType)

		Dim i,j,k,l
		Dim objRS
		Dim intPageCount
		Dim objArticle

		Call CheckParameter(intPage,"int",1)
		Call CheckParameter(intCateId,"int",Empty)
		Call CheckParameter(intAuthorId,"int",Empty)
		Call CheckParameter(dtmYearMonth,"dtm",Empty)

		Title=ZC_BLOG_SUBTITLE

		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn


		'//////////////////////////
		'ontop
		objRS.Source="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Meta] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_ID]>0) AND ([log_Istop]<>0) AND ([log_Level]>1)"
		objRS.Source=objRS.Source & "ORDER BY [log_PostTime] DESC,[log_ID] DESC"
		objRS.Open()
		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = ZC_DISPLAY_COUNT
			intPageCount=objRS.PageCount
			objRS.AbsolutePage = 1

			For i = 1 To objRS.PageSize

				ReDim Preserve aryArticleList(i)

				Set objArticle=New TArticle
				If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16))) Then
					If objArticle.Export(intType)= True Then
						aryArticleList(i)=objArticle.Template_Article_Istop
					End If
				End If
				Set objArticle=Nothing

				objRS.MoveNext
				If objRS.EOF Then Exit For

			Next

		End If
		objRS.Close()
		k=Join(aryArticleList)
		Erase aryArticleList
		'//////////////////////////


		objRS.Source="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Meta] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_ID]>0) AND ([log_Istop]=0) AND ([log_Level]>1)"

		If Not IsEmpty(intCateId) Then
			Dim strSubCateID : strSubCateID=Join(GetSubCateID(intCateId,True),",")
			objRS.Source=objRS.Source & "AND([log_CateID]IN("&strSubCateID&"))"
			If CheckCateByID(intCateId) Then
				Title=TransferHTML(Categorys(intCateId).Name,"[html-format]")
				TemplateTags_ArticleList_Category_ID=Categorys(intCateId).ID
				If Categorys(intCateId).TemplateName<>"" Then
					Categorys(intCateId).html=GetTemplate("TEMPLATE_" & Categorys(intCateId).TemplateName)
					If Categorys(intCateId).html<>"" Then template=Categorys(intCateId).TemplateName
				End If
			End If
		End if
		If Not IsEmpty(intAuthorId) Then
			objRS.Source=objRS.Source & "AND([log_AuthorID]="&intAuthorId&")"
			If CheckAuthorByID(intAuthorId) Then
				Title=TransferHTML(Users(intAuthorId).Name,"[html-format]")
				TemplateTags_ArticleList_Author_ID=Users(intAuthorId).ID
			End If
		End if
		If IsDate(dtmYearMonth) Then
			Dim y
			Dim m
			Dim d
			Dim ny
			Dim nm

			If IsDate(dtmYearMonth) Then
				'dtmYearMonth=CDate(dtmYearMonth)
			Else
				Call showError(3)
			End If

			y=Year(dtmYearMonth)
			m=Month(dtmYearMonth)
			d=Day(dtmYearMonth)

			TemplateTags_ArticleList_Date_ShortDate=dtmYearMonth
			TemplateTags_ArticleList_Date_Year=y
			TemplateTags_ArticleList_Date_Month=m
			TemplateTags_ArticleList_Date_Day=d

			ny=y
			nm=m+1
			If m=12 Then ny=ny+1:nm=1


			If InstrRev(CStr(dtmYearMonth),"-")>=7 Then
				objRS.Source=objRS.Source & "AND(Year([log_PostTime])="&y&") AND(Month([log_PostTime])="&m&") AND(Day([log_PostTime])="&d&")"
			Else
				objRS.Source=objRS.Source & "AND(Year([log_PostTime])="&y&") AND(Month([log_PostTime])="&m&")"
			End If

			Template_Calendar="<script language=""JavaScript"" src="""&ZC_BLOG_HOST&"zb_system/function/c_html_js.asp?date="&dtmYearMonth&""" type=""text/javascript""></script>"

			Title=Year(dtmYearMonth) & " " & ZVA_Month(Month(dtmYearMonth))
		End If
		If Not IsEmpty(strTagsName) Then
			'On Error Resume Next
			Dim Tag
			For Each Tag in Tags
				If IsObject(Tag) Then
					If UCase(Tag.Name)=UCase(strTagsName) Then
						objRS.Source=objRS.Source & "AND([log_Tag] LIKE '%{" & Tag.ID & "}%')"
						TemplateTags_ArticleList_Tags_ID=Tag.ID
						Title=Server.HTMLEncode(strTagsName)
						If Tag.TemplateName<>"" Then
							Tag.html=GetTemplate("TEMPLATE_" & Tag.TemplateName)
							If Tag.html<>"" Then template=Tag.TemplateName
						End If
					End If
				End If
			Next
			'Err.Clear
		End If

		objRS.Source=objRS.Source & "ORDER BY [log_PostTime] DESC,[log_ID] DESC"
		objRS.Open()

		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = ZC_DISPLAY_COUNT
			intPageCount=objRS.PageCount
			objRS.AbsolutePage = intPage

			For i = 1 To objRS.PageSize

				If intPage>intPageCount Then Exit For

				ReDim Preserve aryArticleList(i)

				Set objArticle=New TArticle
				If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16))) Then
					If objArticle.Export(intType)= True Then
						aryArticleList(i)=objArticle.Template_Article_Multi
					End If
				End If
				Set objArticle=Nothing

				objRS.MoveNext
				If objRS.EOF Then Exit For

			Next

		End If

		objRS.Close()
		Set objRS=Nothing

		Template_Article_Multi=k & Join(aryArticleList)

		TemplateTags_ArticleList_Page_Now=intPage
		TemplateTags_ArticleList_Page_All=intPageCount

		Call ExportBar(intPage,intPageCount,intCateId,intAuthorId,dtmYearMonth,strTagsName)

		Export=True

		'plugin node
		bAction_Plugin_TArticleList_Export_End=False
		For Each sAction_Plugin_TArticleList_Export_End in Action_Plugin_TArticleList_Export_End
			If Not IsEmpty(sAction_Plugin_TArticleList_Export_End) Then Call Execute(sAction_Plugin_TArticleList_Export_End)
			If bAction_Plugin_TArticleList_Export_End=True Then Exit Function
		Next

	End Function


	Public Function ExportByCache(intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType)

		'plugin node
		bAction_Plugin_TArticleList_ExportByCache_Begin=False
		For Each sAction_Plugin_TArticleList_ExportByCache_Begin in Action_Plugin_TArticleList_ExportByCache_Begin
			If Not IsEmpty(sAction_Plugin_TArticleList_ExportByCache_Begin) Then Call Execute(sAction_Plugin_TArticleList_ExportByCache_Begin)
			If bAction_Plugin_TArticleList_ExportByCache_Begin=True Then Exit Function
		Next

		'plugin node
		Call Filter_Plugin_TArticleList_ExportByCache(intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType)


		Dim strType
		Dim i,j,s,t,k,l
		Dim intAllPage
		Dim intTagsID
		Dim objArticle

		Call CheckParameter(intPage,"int",1)
		Call CheckParameter(intCateId,"int",Empty)
		Call CheckParameter(intAuthorId,"int",Empty)
		Call CheckParameter(dtmYearMonth,"dtm",Empty)


		i=InStr(1,TagsList,vbTab & strTagsName & vbVerticalTab,vbBinaryCompare)
		If i>0 Then
			j=Left(TagsList,i-1)
			i=InStrRev(j,vbTab)
			intTagsID=Right(j,Len(j)-i)
			Call CheckParameter(intTagsID,"int",Empty)
		End If


		'//////////////////////////
		'ontop
		If True Then
			strType="Istop" & "Page1" & "["
			s="Istop" & "Page"

			i=InStrRev(AllList,s)
			If i>0 Then
				j=InStr(i,AllList,"[",vbBinaryCompare)
				s=Mid(AllList,i+Len(s),j-i-Len(s))
				intAllPage=CInt(s)
			End If

			i=InStr(1,AllList,strType,vbBinaryCompare)
			If i>0 Then
				i=Len(strType)+i
				j=InStr(i,AllList,"]",vbBinaryCompare)
				s=Mid(AllList,i,j-i)
				aryArticle=Split(s,";")
			End If


			If IsArray(aryArticle) Then

				Redim aryArticleList(UBound(aryArticle))

				For i=LBound(aryArticle) To UBound(aryArticle)-1
					Set objArticle = New TArticle
					objArticle.ID=aryArticle(i)
					If objArticle.LoadCache Then
						aryArticleList(i)=objArticle.Template_Article_Multi
					End if
					Set objArticle = Nothing
				Next

				k=Join(aryArticleList)
				Erase aryArticleList
				ReDim aryArticle(0)

			End If

		End If
		'////////////////////////////


		strType="All" & "Page" & CStr(intPage) & "["
		s="All" & "Page"

		Title=ZC_BLOG_SUBTITLE


		i=InStrRev(AllList,s)
		If i>0 Then
			j=InStr(i,AllList,"[",vbBinaryCompare)
			s=Mid(AllList,i+Len(s),j-i-Len(s))
			intAllPage=CInt(s)
		End If

		i=InStr(1,AllList,strType,vbBinaryCompare)
		If i>0 Then
			i=Len(strType)+i
			j=InStr(i,AllList,"]",vbBinaryCompare)
			s=Mid(AllList,i,j-i)
			aryArticle=Split(s,";")
		End If


		If IsArray(aryArticle) Then


			Redim aryArticleList(UBound(aryArticle))

			For i=LBound(aryArticle) To UBound(aryArticle)-1
				Set objArticle = New TArticle
				objArticle.ID=aryArticle(i)
				If objArticle.LoadCache Then
					aryArticleList(i)=objArticle.Template_Article_Multi
				End if
				Set objArticle = Nothing
			Next

			Template_Article_Multi=k & Join(aryArticleList)

		End If

		TemplateTags_ArticleList_Page_Now=intPage
		TemplateTags_ArticleList_Page_All=intAllPage

		Call ExportBar(intPage,intAllPage,intCateId,intAuthorId,dtmYearMonth,strTagsName)

		ExportByCache=True

		'plugin node
		bAction_Plugin_TArticleList_ExportByCache_End=False
		For Each sAction_Plugin_TArticleList_ExportByCache_End in Action_Plugin_TArticleList_ExportByCache_End
			If Not IsEmpty(sAction_Plugin_TArticleList_ExportByCache_End) Then Call Execute(sAction_Plugin_TArticleList_ExportByCache_End)
			If bAction_Plugin_TArticleList_ExportByCache_End=True Then Exit Function
		Next

	End Function


	Public Function ExportByMixed(intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType)

		'plugin node
		bAction_Plugin_TArticleList_ExportByMixed_Begin=False
		For Each sAction_Plugin_TArticleList_ExportByMixed_Begin in Action_Plugin_TArticleList_ExportByMixed_Begin
			If Not IsEmpty(sAction_Plugin_TArticleList_ExportByMixed_Begin) Then Call Execute(sAction_Plugin_TArticleList_ExportByMixed_Begin)
			If bAction_Plugin_TArticleList_ExportByMixed_Begin=True Then Exit Function
		Next

		'plugin node
		Call Filter_Plugin_TArticleList_ExportByMixed(intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType)

		Dim strType
		Dim i,j,k,l,s
		Dim objRS
		Dim intPageCount
		Dim objArticle
		Dim intAllPage

		Call CheckParameter(intPage,"int",1)
		Call CheckParameter(intCateId,"int",Empty)
		Call CheckParameter(intAuthorId,"int",Empty)
		Call CheckParameter(dtmYearMonth,"dtm",Empty)

		Title=ZC_BLOG_SUBTITLE

		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn


		'//////////////////////////
		'ontop
		If True Then
			strType="Istop" & "Page1" & "["
			s="Istop" & "Page"

			i=InStrRev(AllList,s)
			If i>0 Then
				j=InStr(i,AllList,"[",vbBinaryCompare)
				s=Mid(AllList,i+Len(s),j-i-Len(s))
				intAllPage=CInt(s)
			End If

			i=InStr(1,AllList,strType,vbBinaryCompare)
			If i>0 Then
				i=Len(strType)+i
				j=InStr(i,AllList,"]",vbBinaryCompare)
				s=Mid(AllList,i,j-i)
				aryArticle=Split(s,";")
			End If


			If IsArray(aryArticle) Then

				Redim aryArticleList(UBound(aryArticle))

				For i=LBound(aryArticle) To UBound(aryArticle)-1
					Set objArticle = New TArticle
					objArticle.ID=aryArticle(i)
					If objArticle.LoadCache Then
						aryArticleList(i)=objArticle.Template_Article_Multi
					End if
					Set objArticle = Nothing
				Next

				k=Join(aryArticleList)
				Erase aryArticleList
				ReDim aryArticle(0)

			End If

		End If
		'////////////////////////////


		objRS.Source="SELECT [log_ID] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_ID]>0) AND ([log_Istop]=0) AND ([log_Level]>1)"

		If Not IsEmpty(intCateId) Then
			Dim strSubCateID : strSubCateID=Join(GetSubCateID(intCateId,True),",")
			objRS.Source=objRS.Source & "AND([log_CateID]IN("&strSubCateID&"))"
			If CheckCateByID(intCateId) Then
				Title=TransferHTML(Categorys(intCateId).Name,"[html-format]")
				TemplateTags_ArticleList_Category_ID=Categorys(intCateId).ID
				If Categorys(intCateId).TemplateName<>"" Then
					Categorys(intCateId).html=GetTemplate("TEMPLATE_" & Categorys(intCateId).TemplateName)
					If Categorys(intCateId).html<>"" Then template=Categorys(intCateId).TemplateName
				End If
			End If
		End if
		If Not IsEmpty(intAuthorId) Then
			objRS.Source=objRS.Source & "AND([log_AuthorID]="&intAuthorId&")"
			If CheckAuthorByID(intAuthorId) Then
				Title=TransferHTML(Users(intAuthorId).Name,"[html-format]")
				TemplateTags_ArticleList_Author_ID=Users(intAuthorId).ID
			End If
		End If

		If IsDate(dtmYearMonth) Then
			Dim y
			Dim m
			Dim d
			Dim ny
			Dim nm

			If IsDate(dtmYearMonth) Then
			'	dtmYearMonth=CDate(dtmYearMonth)
			Else
				Call showError(3)
			End If

			y=Year(dtmYearMonth)
			m=Month(dtmYearMonth)
			d=Day(dtmYearMonth)

			TemplateTags_ArticleList_Date_ShortDate=dtmYearMonth
			TemplateTags_ArticleList_Date_Year=y
			TemplateTags_ArticleList_Date_Month=m
			TemplateTags_ArticleList_Date_Day=d
			ny=y
			nm=m+1
			If m=12 Then ny=ny+1:nm=1

			If InstrRev(CStr(dtmYearMonth),"-")>=7 Then
				objRS.Source=objRS.Source & "AND(Year([log_PostTime])="&y&") AND(Month([log_PostTime])="&m&") AND(Day([log_PostTime])="&d&")"
			Else
				objRS.Source=objRS.Source & "AND(Year([log_PostTime])="&y&") AND(Month([log_PostTime])="&m&")"
			End If

			Template_Calendar="<script language=""JavaScript"" src="""&ZC_BLOG_HOST&"zb_system/function/c_html_js.asp?date="&dtmYearMonth&""" type=""text/javascript""></script>"

			Title=Year(dtmYearMonth) & " " & ZVA_Month(Month(dtmYearMonth))
		End If

		If Not IsEmpty(strTagsName) Then
				'On Error Resume Next
				Dim Tag
				For Each Tag in Tags
						If IsObject(Tag) Then
								Dim arrTagsName, Tag_i
								arrTagsName = split(strTagsName, ",")
								For Tag_i = 0 To UBound(arrTagsName)
								strTagsName = arrTagsName(Tag_i)
								If UCase(Tag.Name)=UCase(strTagsName) Then
										objRS.Source=objRS.Source & "AND([log_Tag] LIKE '%{" & Tag.ID & "}%')"
										TemplateTags_ArticleList_Tags_ID=Tag.ID
										If Tag.TemplateName<>"" Then
											Tag.html=GetTemplate("TEMPLATE_" & Tag.TemplateName)
											If Tag.html<>"" Then template=Tag.TemplateName
										End If
								End If
								Next'Tag_i
						End If
				Next
				'Err.Clear
				Title=Server.HTMLEncode(strTagsName)

		End If

		objRS.Source=objRS.Source & "ORDER BY [log_PostTime] DESC,[log_ID] DESC"
		objRS.Open()

		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = ZC_DISPLAY_COUNT
			intPageCount=objRS.PageCount
			objRS.AbsolutePage = intPage

			For i = 1 To objRS.PageSize

				If intPage>intPageCount Then Exit For

				ReDim Preserve aryArticleList(i)

				Set objArticle = New TArticle
				objArticle.ID=objRS(0)
				If objArticle.LoadCache Then
					aryArticleList(i)=objArticle.Template_Article_Multi
				End if
				Set objArticle = Nothing

				objRS.MoveNext
				If objRS.EOF Then Exit For

			Next

		End If

		objRS.Close()
		Set objRS=Nothing

		Template_Article_Multi=k & Join(aryArticleList)

		TemplateTags_ArticleList_Page_Now=intPage
		TemplateTags_ArticleList_Page_All=intPageCount

		Call ExportBar(intPage,intPageCount,intCateId,intAuthorId,dtmYearMonth,strTagsName)

		ExportByMixed=True

		'plugin node
		bAction_Plugin_TArticleList_ExportByMixed_End=False
		For Each sAction_Plugin_TArticleList_ExportByMixed_End in Action_Plugin_TArticleList_ExportByMixed_End
			If Not IsEmpty(sAction_Plugin_TArticleList_ExportByMixed_End) Then Call Execute(sAction_Plugin_TArticleList_ExportByMixed_End)
			If bAction_Plugin_TArticleList_ExportByMixed_End=True Then Exit Function
		Next

	End Function



	Public Function Build()

		Dim aryTemplateTagsName
		Dim aryTemplateTagsValue

		Dim aryTemplateSubName()
		Dim aryTemplateSubValue()

		Dim i,j

		If IsEmpty(html) Then html=template

		'plugin node
		Call Filter_Plugin_TArticleList_Build_Template(html)

		ReDim aryTemplateSubName(14)
		ReDim aryTemplateSubValue(14)

		aryTemplateSubName(  1)="template:article-multi"
		aryTemplateSubValue( 1)=Template_Article_Multi
		aryTemplateSubName(  2)="template:pagebar"
		aryTemplateSubValue( 2)=Template_PageBar
		aryTemplateSubName(  3)="template:pagebar_next"
		aryTemplateSubValue( 3)=Template_PageBar_Next
		aryTemplateSubName(  4)="template:pagebar_previous"
		aryTemplateSubValue( 4)=Template_PageBar_Previous
		aryTemplateSubName(  5)="articlelist/author/id"
		aryTemplateSubValue( 5)=TemplateTags_ArticleList_Author_ID
		aryTemplateSubName(  6)="articlelist/tags/id"
		aryTemplateSubValue( 6)=TemplateTags_ArticleList_Tags_ID
		aryTemplateSubName(  7)="articlelist/category/id"
		aryTemplateSubValue( 7)=TemplateTags_ArticleList_Category_ID
		aryTemplateSubName(  8)="articlelist/date/year"
		aryTemplateSubValue( 8)=TemplateTags_ArticleList_Date_Year
		aryTemplateSubName(  9)="articlelist/date/month"
		aryTemplateSubValue( 9)=TemplateTags_ArticleList_Date_Month
		aryTemplateSubName( 10)="articlelist/date/day"
		aryTemplateSubValue(10)=TemplateTags_ArticleList_Date_Day
		aryTemplateSubName( 11)="articlelist/date/shortdate"
		aryTemplateSubValue(11)=TemplateTags_ArticleList_Date_ShortDate
		aryTemplateSubName( 12)="articlelist/page/now"
		aryTemplateSubValue(12)=TemplateTags_ArticleList_Page_Now
		aryTemplateSubName( 13)="articlelist/page/all"
		aryTemplateSubValue(13)=TemplateTags_ArticleList_Page_All
		aryTemplateSubName( 14)="articlelist/page/count"
		aryTemplateSubValue(14)=ZC_DISPLAY_COUNT

		'plugin node
		Call Filter_Plugin_TArticleList_Build_TemplateSub(aryTemplateSubName,aryTemplateSubValue)

		j=UBound(aryTemplateSubName)
		For i=0 to j
			html=Replace(html,"<#" & aryTemplateSubName(i) & "#>",aryTemplateSubValue(i))
		Next

		aryTemplateTagsName=TemplateTagsName
		aryTemplateTagsValue=TemplateTagsValue

		aryTemplateTagsName(0)="BlogTitle"
		aryTemplateTagsValue(0)=Title

		Call Filter_Plugin_TArticleList_Build_TemplateTags(aryTemplateTagsName,aryTemplateTagsValue)

		j=UBound(aryTemplateTagsName)
		For i=1 to j
			html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
		Next
		html=Replace(html,"<#" & aryTemplateTagsName(0) & "#>",aryTemplateTagsValue(0))

		If IsEmpty(Template_Calendar) Or Len(Template_Calendar)=0 Then
			For i=1 to j
				If aryTemplateTagsName(i)="CACHE_INCLUDE_CALENDAR" Then
					Template_Calendar=aryTemplateTagsValue(i)
				End If
			Next
		End If
		html=Replace(html,"<#CACHE_INCLUDE_CALENDAR_NOW#>",Template_Calendar)

		Build=True

	End Function


	Function Save()

		html=TransferHTML(html,"[no-asp]")
		If ZC_STATIC_TYPE="asp" Then
			html="<"&"%@ CODEPAGE=65001 %"&">" & html
		End If

		Call SaveToFile(BlogPath & Directory & FileName,html,"utf-8",False)

		Save=True

	End Function


	Function SetVar(TemplateTag,TemplateValue)

		If IsEmpty(html) Then html=template

		html=Replace(html,"<#" & TemplateTag & "#>",TemplateValue)

	End Function


	Public Function Search(strQuestion)

		'plugin node
		bAction_Plugin_TArticleList_Search_Begin=False
		For Each sAction_Plugin_TArticleList_Search_Begin in Action_Plugin_TArticleList_Search_Begin
			If Not IsEmpty(sAction_Plugin_TArticleList_Search_Begin) Then Call Execute(sAction_Plugin_TArticleList_Search_Begin)
			If bAction_Plugin_TArticleList_Search_Begin=True Then Exit Function
		Next

		Dim i
		Dim j
		Dim s

		Dim objRS
		Dim intPageCount
		Dim objArticle

		strQuestion=Trim(strQuestion)

		If Len(strQuestion)=0 Then Search=True:Exit Function
		'If CheckRegExp(strQuestion,"[nojapan]") Then Exit Function

		strQuestion=FilterSQL(strQuestion)

		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn




		objRS.Source="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Meta] FROM [blog_Article] WHERE ([log_CateID]>0) And ([log_ID]>0) AND ([log_Level]>2)"

		If ZC_MSSQL=False Then
			objRS.Source=objRS.Source & "AND( (InStr(1,LCase([log_Title]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('"&strQuestion&"'),0)<>0) )"
		Else
			objRS.Source=objRS.Source & "AND( (CHARINDEX('"&strQuestion&"',[log_Title])<>0) OR (CHARINDEX('"&strQuestion&"',[log_Intro])<>0) OR (CHARINDEX('"&strQuestion&"',[log_Content])<>0) )"
		End If

		objRS.Source=objRS.Source & "ORDER BY [log_PostTime] DESC,[log_ID] DESC"

		objRS.Open()

		's=Replace(Replace(ZC_MSG086,"%s","<strong>" & TransferHTML(Replace(strQuestion,Chr(39)&Chr(39),Chr(39)),"[html-format]") & "</strong>",vbTextCompare,1),"%s","<strong>" & objRS.RecordCount & "</strong>")
		s=Replace(Replace(ZC_MSG086,"%s","<strong>" & TransferHTML(Replace(strQuestion,Chr(39)&Chr(39),Chr(39),1,-1,0),"[html-format]") & "</strong>",vbTextCompare,1),"%s","<strong>" & objRS.RecordCount & "</strong>",1,-1,0)

		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = ZC_SEARCH_COUNT
			intPageCount=objRS.PageCount
			objRS.AbsolutePage = 1

			For i = 1 To objRS.PageSize

				ReDim Preserve aryArticleList(i)

				Set objArticle=New TArticle
				If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16))) Then
					If objArticle.Export(ZC_DISPLAY_MODE_SEARCH)= True Then
						aryArticleList(i)=objArticle.Template_Article_Search
					End If
				End If
				Set objArticle=Nothing

				objRS.MoveNext
				If objRS.EOF Then Exit For

			Next

		End If

		objRS.Close()
		Set objRS=Nothing

		Template_Article_Multi=Join(aryArticleList)

		Title=TransferHTML(strQuestion,"[html-format]")

		Search=True

		'plugin node
		bAction_Plugin_TArticleList_Search_End=False
		For Each sAction_Plugin_TArticleList_Search_End in Action_Plugin_TArticleList_Search_End
			If Not IsEmpty(sAction_Plugin_TArticleList_Search_End) Then Call Execute(sAction_Plugin_TArticleList_Search_End)
			If bAction_Plugin_TArticleList_Search_End=True Then Exit Function
		Next

	End Function


	Public Function ExportBar(intNowPage,intAllPage,intCateId,intAuthorId,dtmYearMonth,strTagsName)

		'plugin node
		bAction_Plugin_TArticleList_ExportBar_Begin=False
		For Each sAction_Plugin_TArticleList_ExportBar_Begin in Action_Plugin_TArticleList_ExportBar_Begin
			If Not IsEmpty(sAction_Plugin_TArticleList_ExportBar_Begin) Then Call Execute(sAction_Plugin_TArticleList_ExportBar_Begin)
			If bAction_Plugin_TArticleList_ExportBar_Begin=True Then Exit Function
		Next

		Dim i
		Dim s
		Dim t
		Dim strPageBar

		If Not IsEmpty(intCateId) Then t=t & "cate=" & intCateId & "&amp;"
		If Not IsEmpty(dtmYearMonth) Then
			t=t & "date=" & Year(dtmYearMonth) & "-" & Month(dtmYearMonth)
			If InstrRev(CStr(dtmYearMonth),"-")>=7 Then
				t=t & "-" &  Day(dtmYearMonth)
			End If
			t=t & "&amp;"
		End If
		If Not IsEmpty(intAuthorId) Then t=t & "auth=" & intAuthorId & "&amp;"
		If Not (strTagsName="") Then t=t & "tags=" & Server.URLEncode(strTagsName) & "&amp;"
		If intAllPage>0 Then
			Dim a,b

			s=ZC_BLOG_HOST & "catalog.asp?"& t &"page=1"

			strPageBar=GetTemplate("TEMPLATE_B_PAGEBAR")
			strPageBar=Replace(strPageBar,"<#pagebar/page/url#>",s)
			strPageBar=Replace(strPageBar,"<#pagebar/page/number#>",ZC_MSG285)
			Template_PageBar=Template_PageBar & strPageBar

			If intAllPage>ZC_PAGEBAR_COUNT Then
				a=intNowPage
				b=intNowPage+ZC_PAGEBAR_COUNT
				If a>ZC_PAGEBAR_COUNT Then a=a-1:b=b-1
				If b>intAllPage Then b=intAllPage:a=intAllPage-ZC_PAGEBAR_COUNT
			Else
				a=1:b=intAllPage
			End If
			For i=a to b

				s=ZC_BLOG_HOST & "catalog.asp?"& t &"page="& i

				strPageBar=GetTemplate("TEMPLATE_B_PAGEBAR")
				If i=intNowPage then
					Template_PageBar=Template_PageBar & "<span class=""now-page"">" & i & "</span>"
				Else
					strPageBar=Replace(strPageBar,"<#pagebar/page/url#>",s)
					strPageBar=Replace(strPageBar,"<#pagebar/page/number#>",i)
					Template_PageBar=Template_PageBar & strPageBar
				End If

			Next

			s=ZC_BLOG_HOST & "catalog.asp?"& t &"page="& intAllPage

			strPageBar=GetTemplate("TEMPLATE_B_PAGEBAR")
			strPageBar=Replace(strPageBar,"<#pagebar/page/url#>",s)
			strPageBar=Replace(strPageBar,"<#pagebar/page/number#>",ZC_MSG286)
			Template_PageBar=Template_PageBar & strPageBar

			If intNowPage=1 Then
				Template_PageBar_Previous=""
			Else
				Template_PageBar_Previous="<span class=""pagebar-previous""><a href="""& ZC_BLOG_HOST &"catalog.asp?"& t &"page="& intNowPage-1 &"""><span>"&ZC_MSG156&"</span></a></span>"

			End If

			If intNowPage=intAllPage Then
				Template_PageBar_Next=""
			Else
				Template_PageBar_Next="<span class=""pagebar-next""><a href="""& ZC_BLOG_HOST &"catalog.asp?"& t &"page="& intNowPage+1 &"""><span>"&ZC_MSG155&"</span></a></span>"
			End If

		End If

		ExportBar=True

		'plugin node
		bAction_Plugin_TArticleList_ExportBar_End=False
		For Each sAction_Plugin_TArticleList_ExportBar_End in Action_Plugin_TArticleList_ExportBar_End
			If Not IsEmpty(sAction_Plugin_TArticleList_ExportBar_End) Then Call Execute(sAction_Plugin_TArticleList_ExportBar_End)
			If bAction_Plugin_TArticleList_ExportBar_End=True Then Exit Function
		Next

	End Function


	Public Function LoadCache()

		Dim strContent

		strContent=""
		strContent=LoadFromFile(BlogPath & "zb_users/CACHE/cache_list_"&ZC_BLOG_CLSID&".html","utf-8")
		AllList=strContent

		LoadCache=True

	End Function


	Private Sub Class_Initialize()

		Redim Article(ZC_DISPLAY_COUNT)

	End Sub

End Class
'*********************************************************




'*********************************************************
' 目的：    定义TUser类
' 输入：    无
' 返回：    无
'*********************************************************
Class TUser

	Public ID
	Public Level
	Public Name
	Public Password
	Public Alias

	Public Sex
	Public Email
	Public MSN
	Public QQ
	Public HomePage
	Public Intro

	Public Count

	Public LastVisitTime
	Public LastVisitIP
	Public Meta

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property

	Public html

	Public Property Get Url

		'plugin node
		bAction_Plugin_TUser_Url=False
		For Each sAction_Plugin_TUser_Url in Action_Plugin_TUser_Url
			If Not IsEmpty(sAction_Plugin_TUser_Url) Then Call Execute(sAction_Plugin_TUser_Url)
			If bAction_Plugin_TUser_Url=True Then Exit Property
		Next

		Url = ZC_BLOG_HOST & "catalog.asp?"& "auth=" & ID

		'If CheckPluginState("STACentre") Then
		'	'Dim objAuth : Set objAuth=New STACentre_Authors : If objAuth.LoadInfoByID(ID) Then Url=objAuth.Url : Set objAuth=Nothing
		'End If

		Call Filter_Plugin_TUser_Url(Url)

	End Property

	Public Property Get HtmlUrl
		HtmlUrl=TransferHTML(Url,"[html-format]")
	End Property

	Public Property Get RssUrl
		RssUrl = ZC_BLOG_HOST & "sydication.asp?user=" & ID
	End Property

	Private FLoginType
	Public Property Let LoginType(strLoginType)
			FLoginType=strLoginType
	End Property
	Public Property Get LoginType
			LoginType = FLoginType
	End Property


	Public Function GetMixPassword(md5password)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [mem_Guid] FROM [blog_Member] WHERE [mem_Name]='"&Name & "'" )
		If (Not objRS.Bof) And (Not objRS.Eof) Then
			GetMixPassword=MD5(md5password & objRS("mem_Guid"))
		End If

		objRS.Close
		Set objRS=Nothing

	End Function


	Public Function Verify()
		Dim strUserName
		Dim  strPassWord

		If LoginType="Cookies" Then
			strPassWord=Request.Cookies("password")
			If (strPassWord="") Then Exit Function
			strUserName=vbsunescape(Request.Cookies("username"))
			If (strUserName="") Then Exit Function
		ElseIf LoginType="Form" Then
			strPassWord=Request.Form("password")
			If (strPassWord="") Then Exit Function
			strUserName=Request.Form("username")
			If (strUserName="") Then Exit Function
		ElseIf LoginType="QueryString" Then
			strPassWord=Request.QueryString("password")
			If (strPassWord="") Then Exit Function
			strUserName=Request.QueryString("username")
			If (strUserName="") Then Exit Function
		ElseIf LoginType="Self" Then
			strPassWord=Password
			If (strPassWord="") Then Exit Function
			strUserName=Name
			If (strUserName="") Then Exit Function
		Else
			Exit Function
		End If

		strUserName=FilterSQL(strUserName)
		strPassWord=FilterSQL(strPassWord)

		'校检
		If Len(strUserName) >ZC_USERNAME_MAX Then Call ShowError(7)
		If Len(strPassWord)<>ZC_PASSWORD_MAX Then Call ShowError(7)
		If Not CheckRegExp(strUserName,"[username]") Then Call ShowError(7)
		If Not CheckRegExp(strPassWord,"[password]") Then Call ShowError(7)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [mem_ID],[mem_Level],[mem_Password],[mem_Guid] FROM [blog_Member] WHERE [mem_Name]='"&strUserName & "'" )
		If (Not objRS.Bof) And (Not objRS.Eof) Then

			If LoginType<>"Cookies" Then
				strPassWord=MD5(strPassWord & objRS("mem_Guid"))
			End If

			If StrComp(strPassWord,objRS("mem_Password"))=0 Then

				ID=objRS("mem_ID")
				LoadInfobyID(ID)
				Verify=True

			Else
				'If LoginType="Cookies" Then Response.Cookies("password")=""
			End If
		Else
			'If LoginType="Cookies" Then Response.Cookies("password")=""
		End If

		objRS.Close
		Set objRS=Nothing

	End Function


	Function LoadInfobyID(user_ID)

		Call CheckParameter(user_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [mem_ID],[mem_Name],[mem_Level],[mem_Password],[mem_Email],[mem_HomePage],[mem_PostLogs],[mem_Intro],[mem_Meta] FROM [blog_Member] WHERE [mem_ID]=" & user_ID)
		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("mem_ID")
			Name=objRS("mem_Name")
			Level=objRS("mem_Level")
			Password=objRS("mem_Password")
			Email=objRS("mem_Email")
			HomePage=objRS("mem_HomePage")
			Count=objRS("mem_PostLogs")
			Alias=objRS("mem_Intro")
			
			MetaString=objRS("mem_Meta")

			If IsNull(Email) Or IsEmpty(Email) Or Len(Email)=0 Then Email="a@b.com"
			If IsNull(HomePage) Then HomePage=""
			If IsNull(Alias) Then Alias=""

			LoadInfobyID=True
		End If
		objRS.Close
		Set objRS=Nothing


		Call Filter_Plugin_TUser_LoadInfobyID(ID,Name,Level,Password,Email,HomePage,Count,Alias,MetaString)

	End Function


	Public Function LoadInfoByArray(aryUserInfo)

		If IsArray(aryUserInfo)=True Then

			ID=aryUserInfo(0)
			Name=aryUserInfo(1)
			Level=aryUserInfo(2)
			Password=aryUserInfo(3)
			Email=aryUserInfo(4)
			HomePage=aryUserInfo(5)
			Count=aryUserInfo(6)
			Alias=aryUserInfo(7)
			MetaString=aryUserInfo(8)

		End If

		If IsNull(Email) Or IsEmpty(Email) Or Len(Email)=0 Then Email="a@b.com"
		If IsNull(HomePage) Then HomePage=""
		If IsNull(Alias) Then Alias=""

		LoadInfoByArray=True

		Call Filter_Plugin_TUser_LoadInfoByArray(ID,Name,Level,Password,Email,HomePage,Count,Alias,MetaString)

	End Function


	Function Edit(currentUser)

		Call Filter_Plugin_TUser_Edit(ID,Name,Level,Password,Email,HomePage,Count,Alias,MetaString,currentUser)

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(Level,"int",0)
		Dim Guid
		Guid=RndGuid()
		If ((Level<1) Or (Level>5)) Then Call ShowError(16)
		If (Name="") Then Call ShowError(7)
		If Len(Name) >ZC_USERNAME_MAX Then Call ShowError(7)
		If Not CheckRegExp(Name,"[username]") Then Call ShowError(7)

		Email=FilterSQL(Email)
		HomePage=FilterSQL(HomePage)

		Email=TransferHTML(Email,"[html-format]")
		HomePage=TransferHTML(HomePage,"[html-format]")

		Alias=TransferHTML(Alias,"[filename]")
		Alias=FilterSQL(Alias)

		If Len(Email)=0 Then Call ShowError(29)
		If Len(Email)>ZC_EMAIL_MAX Then Call ShowError(29)
		If Len(HomePage)>ZC_HOMEPAGE_MAX Then Call ShowError(29)

		If Not CheckRegExp(Email,"[email]") Then Call ShowError(29)
		IF Len(HomePage)>0 Then
			If Not CheckRegExp(HomePage,"[homepage]") Then Call ShowError(30)
		End If
		PassWord=MD5(PassWord & Guid)

		If ID=0 Then

			If Level <= currentUser.Level Then ShowError(6)
			If Len(PassWord)<>ZC_PASSWORD_MAX Then Call ShowError(7)
			If Not CheckRegExp(PassWord,"[password]") Then Call ShowError(7)
			objConn.Execute("INSERT INTO [blog_Member]([mem_Level],[mem_Name],[mem_PassWord],[mem_Email],[mem_HomePage],[mem_Intro],[mem_Guid],[mem_Meta]]) VALUES ("&Level&",'"&Name&"','"&PassWord&"','"&Email&"','"&HomePage&"','"&Alias&"','"&Guid&"','"&MetaString&"')")
			
			Dim objRS
			Set objRS=objConn.Execute("SELECT MAX([mem_ID]) FROM [blog_Member]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
			End If
			Set objRS=Nothing

			Edit=True

		Else

			If (ID=currentUser.ID) And (Level <> currentUser.Level) Then ShowError(6)
			If (ID<>currentUser.ID) And (Level <= currentUser.Level) Then ShowError(6)

			Dim targetUser
			Set targetUser=New TUser
			If targetUser.LoadInfobyID(ID) Then

				If Len(PassWord)=0 Then PassWord=targetUser.PassWord
				If Len(PassWord)<>ZC_PASSWORD_MAX Then Call ShowError(6)
				If Not CheckRegExp(PassWord,"[password]") Then Call ShowError(7)

			Else
				Exit Function
			End If

			objConn.Execute("UPDATE [blog_Member] SET [mem_Level]="&Level&",[mem_Name]='"&Name&"',[mem_PassWord]='"&PassWord&"',[mem_Email]='"&Email&"',[mem_HomePage]='"&HomePage&"',[mem_Intro]='"&Alias&"',[mem_Guid]='"&Guid&"',[mem_Meta]='"&MetaString&"' WHERE [mem_ID]="&ID)

			If Name <> targetUser.Name Then
				objConn.Execute("UPDATE [blog_Comment] SET [comm_Author]='"&Name&"' WHERE [comm_AuthorID]="&ID)
			End If
			If Email <> targetUser.Email Then
				objConn.Execute("UPDATE [blog_Comment] SET [comm_Email]='"&Email&"' WHERE [comm_AuthorID]="&ID)
			End If

			Edit=True

		End If

	End Function


	Function Register(currentUser)

		Call Filter_Plugin_TUser_Register(ID,Name,Level,Password,Email,HomePage,Count,Alias,MetaString,currentUser)

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(Level,"int",0)

		Dim Guid
		Guid=RndGuid()
		If (Level<>4) Then Call ShowError(16)
		If (Name="") Then Call ShowError(7)
		If Len(Name) >ZC_USERNAME_MAX Then Call ShowError(7)
		If Not CheckRegExp(Name,"[username]") Then Call ShowError(7)

		Email=FilterSQL(Email)
		HomePage=FilterSQL(HomePage)

		Email=TransferHTML(Email,"[html-format]")
		HomePage=TransferHTML(HomePage,"[html-format]")

		Alias=TransferHTML(Alias,"[filename]")
		Alias=FilterSQL(Alias)

		If Len(Email)=0 Then Call ShowError(29)
		If Len(Email)>ZC_EMAIL_MAX Then Call ShowError(29)
		If Len(HomePage)>ZC_HOMEPAGE_MAX Then Call ShowError(29)

		If Not CheckRegExp(Email,"[email]") Then Call ShowError(30)
		IF Len(HomePage)>0 Then
			If Not CheckRegExp(HomePage,"[homepage]") Then Call ShowError(30)
		End If

		PassWord=MD5(Password & Guid)
		If ID=0 Then

			If Level <= 1 Then ShowError(6)
			If Len(PassWord)<>ZC_PASSWORD_MAX Then Call ShowError(7)
			If Not CheckRegExp(PassWord,"[password]") Then Call ShowError(7)

			objConn.Execute("INSERT INTO [blog_Member]([mem_Level],[mem_Name],[mem_PassWord],[mem_Email],[mem_HomePage],[mem_Intro],[mem_Guid],[mem_Meta]) VALUES ("&Level&",'"&Name&"','"&PassWord&"','"&Email&"','"&HomePage&"','"&Alias&"','"&Guid&"','"&MetaString&"')")

			Dim objRS
			Set objRS=objConn.Execute("SELECT MAX([mem_ID]) FROM [blog_Member]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
			End If
			Set objRS=Nothing

			Register=True

		End If

	End Function


	Function Del(currentUser)

		Call Filter_Plugin_TUser_Del(ID,Name,Level,Password,Email,HomePage,Count,Alias,MetaString,currentUser)

		Dim objRS
		Dim objUpLoadFile

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(Level,"int",0)

		Dim targetUser
		Set targetUser=New TUser
		If targetUser.LoadInfobyID(ID) Then
			If targetUser.Level<= currentUser.Level Then ShowError(6)
			If currentUser.ID = targetUser.ID Then ShowError(17)
		Else
			Exit Function
		End If

		objConn.Execute("DELETE FROM [blog_Article] WHERE [log_AuthorID] =" & ID)
		objConn.Execute("DELETE FROM [blog_Comment] WHERE [comm_AuthorID] =" & ID)
		objConn.Execute("DELETE FROM [blog_Member] WHERE [mem_ID] =" & ID)

		Set objRS=objConn.Execute("SELECT * FROM [blog_UpLoad] WHERE [ul_AuthorID] =" & ID)
		If (Not objRS.bof) And (Not objRS.eof) Then
			Do While Not objRS.eof
				Set objUpLoadFile=New TUpLoadFile
				If objUpLoadFile.LoadInfoByID(objRS("ul_ID")) Then objUpLoadFile.Del
				Set objUpLoadFile=Nothing
				objRS.MoveNext
			Loop
		End If
		objRS.Close
		Set objRS=Nothing

		objConn.Execute("DELETE FROM [blog_UpLoad] WHERE [ul_AuthorID] =" & ID)

		Del=True

	End Function


	Private Sub Class_Initialize()

		Level=5
		ID=0
		Name=ZC_MSG018

		Sex=0
		Email=""
		MSN=""
		QQ=""
		HomePage=""
		Intro=""

		LoginType="Cookies"

		Set Meta=New TMeta

	End Sub


End Class
'*********************************************************


'为了 Ctrl+F方便。。
'Mark1 By ZSXSOFT  
'*********************************************************
' 目的：    定义TComment类
' 输入：    无
' 返回：    无
'*********************************************************
Class TComment

	Public ID
	Public log_ID
	Public ParentID
	Public ParentCount
	
	Public AuthorID
	Public Author
	Public Content
	Public Email
	Public HomePage

	Public PostTime
	Public IP
	Public Agent

	Public Reply
	Public LastReplyIP
	Public LastReplyTime

	Public Count
	Public Meta

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property

	Public html

	Public Property Get HomePageForAntiSpam
		HomePageForAntiSpam=URLEncodeForAntiSpam(HomePage)
	End Property

	Public Property Get SafeEmail
		If (Email="") Or IsEmpty(Email) Or IsNull(Email) Then Email="null@null.com"
		SafeEmail=Replace(Email,"@","[AT]")
	End Property

	Public Property Get EmailMD5

		If AuthorID>0 Then
			EmailMD5=MD5(Users(AuthorID).Email)
		Else
			If (Email="") Or IsEmpty(Email) Or IsNull(Email) Then
				EmailMD5=""
			Else
				EmailMD5=MD5(Email)
			End If
		End If


	End Property

	Public Property Get FirstContact
		If Len(HomePage)>0 Then
			FirstContact=HomePageForAntiSpam
		Else
			If (Email="") Or IsEmpty(Email) Or IsNull(Email) Then
				FirstContact=""
			Else
				FirstContact="mailto:" & SafeEmail
			End If
		End If
	End Property

	Public Property Get HtmlContent
		'HtmlContent=TransferHTML(UBBCode(Content,"[font][face]"),"[enter][nofollow]")
		HtmlContent=TransferHTML(UBBCode(Content & Reply,"[link][link-antispam][font][face]"),"[enter][nofollow]")
	End Property


	Public Function Post()

		Call Filter_Plugin_TComment_Post(ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,ParentCount,MetaString)
		If IP="" Then
			IP=Request.ServerVariables("REMOTE_ADDR")
			Agent=Request.ServerVariables("HTTP_USER_AGENT")
		End If
		If Len(HomePage)>0 Then
			If InStr(HomePage,"http://")=0 Then HomePage="http://" & HomePage
		End If

		'检查参数
		Call CheckParameter(log_ID,"int",0)
		Call CheckParameter(AuthorID,"int",0)
		Call CheckParameter(PostTime,"dtm",GetTime(Now()))
		Call CheckParameter(ParentID,"int",0)
		Call CheckParameter(ParentCount,"int",0)
		If ParentCount>ZC_MAXFLOOR Then Exit Function
		If ParentID="" Then ParentID=0
		Author=FilterSQL(Author)
		Content=FilterSQL(Content)
		Email=FilterSQL(Email)
		HomePage=FilterSQL(HomePage)

		PostTime=FilterSQL(PostTime)
		IP=FilterSQL(IP)
		Agent=FilterSQL(Agent)

		Reply=FilterSQL(Reply)
		LastReplyIP=FilterSQL(LastReplyIP)

		'作者不能为空
		If Len(Author)=0 Then Call  ShowError(15)
		'If Len(Content)=0 Then Call  ShowError(46)
		'If Len(Content)>ZC_CONTENT_MAX Then Call  ShowError(46)

		Author=TransferHTML(Author,"[html-format]")
		Email=TransferHTML(Email,"[html-format]")
		HomePage=TransferHTML(HomePage,"[html-format]")
		Content=TransferHTML(Content,"[html-format]")

		If Len(Author)>ZC_USERNAME_MAX Then Call  ShowError(15)
		If Len(Email)>ZC_EMAIL_MAX Then Call  ShowError(29)
		If Len(HomePage)>ZC_HOMEPAGE_MAX Then Call  ShowError(30)

		If Not CheckRegExp(Author,"[username]") Then Call  ShowError(15)

		IF Len(Email)>0 Then
			If Not CheckRegExp(Email,"[email]") Then Call  ShowError(29)
		End If

		IF Len(HomePage)>0 Then
			If Not CheckRegExp(HomePage,"[homepage]") Then Call  ShowError(30)
		End If

		Dim objRS
		Dim strSpamIP
		Dim strSpamContent

		Set objRS=objConn.Execute("SELECT [comm_IP],[comm_Content] FROM [blog_Comment] WHERE [comm_ID]= ( SELECT MAX(comm_ID) FROM [blog_Comment] )")

		If (Not objRS.bof) And (Not objRS.eof) Then
			strSpamIP=objRS("comm_IP")
			strSpamContent=objRS("comm_Content")
		End If

		If IsDate(LastReplyTime)=False Then LastReplyTime=GetTime(Now())

		objRS.Close
		Set objRS=Nothing

		If (ID=0) And (strSpamIP=IP) And (strSpamContent=Content) Then
			Call ShowError(39)
		End If

		If ID=0 Then
			objConn.Execute("INSERT INTO [blog_Comment]([log_ID],[comm_AuthorID],[comm_Author],[comm_Content],[comm_Email],[comm_HomePage],[comm_IP],[comm_PostTime],[comm_Agent],[comm_Reply],[comm_LastReplyIP],[comm_LastReplyTime],[comm_ParentID],[comm_ParentCount],[comm_Meta]) VALUES ("&log_ID&","&AuthorID&",'"&Author&"','"&Content&"','"&Email&"','"&HomePage&"','"&IP&"','"&PostTime&"','"&Agent&"','"&Reply&"','"&LastReplyIP&"','"&LastReplyTime&"','"&ParentID&"','"&ParentCount&"','"&MetaString&"')")
			Set objRS=objConn.Execute("SELECT MAX([comm_ID]) FROM [blog_Comment]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
			End If
			Set objRS=Nothing
		Else
			objConn.Execute("UPDATE [blog_Comment] SET [log_ID]="&log_ID&", [comm_AuthorID]="&AuthorID&",[comm_Author]='"&Author&"',[comm_Content]='"&Content&"',[comm_Email]='"&Email&"',[comm_HomePage]='"&HomePage&"',[comm_IP]='"&IP&"',[comm_PostTime]='"&PostTime&"',[comm_Agent]='"&Agent&"',[comm_Reply]='"&Reply&"',[comm_LastReplyIP]='"&LastReplyIP&"',[comm_LastReplyTime]='"&LastReplyTime&"',[comm_ParentID]='"&ParentID&"',[comm_Meta]='"&MetaString&"' WHERE [comm_ID] =" & ID)
		End If

		Post=True

	End Function


	Public Function Del()

		Call Filter_Plugin_TComment_Del(ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,ParentCount,MetaString)

		Call CheckParameter(ID,"int",0)
		If (ID=0) Then Del=False:Exit Function
		objConn.Execute("DELETE FROM [blog_Comment] WHERE [comm_ID] =" & ID)
		Del=True
	End Function


	Public Function LoadInfoByID(comm_ID)

		Call CheckParameter(comm_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [comm_ID],[log_ID],[comm_AuthorID],[comm_Author],[comm_Content],[comm_Email],[comm_HomePage],[comm_PostTime],[comm_IP],[comm_Agent],[comm_Reply],[comm_LastReplyIP],[comm_LastReplyTime],[comm_ParentID],[comm_ParentCount],[comm_Meta] FROM [blog_Comment] WHERE [comm_ID]=" & comm_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("comm_ID")
			log_ID=objRS("log_ID")
			AuthorID=objRS("comm_AuthorID")
			Author=objRS("comm_Author")
			Content=objRS("comm_Content")
			Email=objRS("comm_Email")
			HomePage=objRS("comm_HomePage")
			PostTime=objRS("comm_PostTime")
			IP=objRS("comm_IP")
			Agent=objRS("comm_Agent")
			Reply=objRS("comm_Reply")
			LastReplyIP=objRS("comm_LastReplyIP")
			LastReplyTime=objRS("comm_LastReplyTime")
			ParentID=objRS("comm_ParentID")
			ParentCount=objRS("comm_ParentCount")
			MetaString=objRS("comm_Meta")

			LoadInfoByID=True

		End If

		objRS.Close
		Set objRS=Nothing

		If IsNull(HomePage) Then HomePage=""

		Call Filter_Plugin_TComment_LoadInfoByID(ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,ParentCount,MetaString)

	End Function


	Public Function LoadInfoByArray(aryCommInfo)

		If IsArray(aryCommInfo)=True Then
			ID=aryCommInfo(0)
			log_ID=aryCommInfo(1)
			AuthorID=aryCommInfo(2)
			Author=aryCommInfo(3)
			Content=aryCommInfo(4)
			Email=aryCommInfo(5)
			HomePage=aryCommInfo(6)
			PostTime=aryCommInfo(7)
			IP=aryCommInfo(8)
			Agent=aryCommInfo(9)
			Reply=aryCommInfo(10)
			LastReplyIP=aryCommInfo(11)
			LastReplyTime=aryCommInfo(12)
			ParentID=aryCommInfo(13)
			ParentCount=aryCommInfo(14)
			MetaString=aryCommInfo(15)

		End If

		If IsNull(HomePage) Then HomePage=""

		LoadInfoByArray=True

		Call Filter_Plugin_TComment_LoadInfoByArray(ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,ParentCount,MetaString)

	End Function

'Mark2 MakeTemplate
	Public Function MakeTemplate(strC,isChild)
		Dim html,i,j,template
		html=strC
		template=GetTemplate("TEMPLATE_B_ARTICLE_COMMENTREV")
		if template="" then template=strc
		'plugin node
		Call Filter_Plugin_TComment_MakeTemplate_Template(html)
		Dim ChildTemplate
		Dim aryTemplateTagsName()
		Dim aryTemplateTagsValue()
		ReDim aryTemplateTagsName(13)
		ReDim aryTemplateTagsValue(13)
		If ParentID="" Then ParentID=0
		If ParentID>0 Then
			If isChild=False Then MakeTemplate="":Exit Function
		End If
		aryTemplateTagsName(  1)="article/comment/id"
		aryTemplateTagsValue( 1)=ID
		aryTemplateTagsName(  2)="article/comment/name"
		aryTemplateTagsValue( 2)=Author
		aryTemplateTagsName(  3)="article/comment/url"
		aryTemplateTagsValue( 3)=HomePage
		aryTemplateTagsName(  4)="article/comment/urlencoder"
		aryTemplateTagsValue( 4)=HomePageForAntiSpam
		aryTemplateTagsName(  5)="article/comment/email"
		aryTemplateTagsValue( 5)=SafeEmail
		aryTemplateTagsName(  6)="article/comment/posttime"
		aryTemplateTagsValue( 6)=PostTime
		aryTemplateTagsName(  7)="article/comment/content"
		aryTemplateTagsValue( 7)=HtmlContent
		aryTemplateTagsName(  8)="article/comment/count"
		aryTemplateTagsValue( 8)=Count
		aryTemplateTagsName(  9)="article/comment/authorid"
		aryTemplateTagsValue( 9)=AuthorID
		aryTemplateTagsName( 10)="article/comment/firstcontact"
		aryTemplateTagsValue(10)=FirstContact
		aryTemplateTagsName( 11)="article/comment/emailmd5"
		aryTemplateTagsValue(11)=EmailMD5
		aryTemplateTagsName( 12)="article/comment/parentid"
		aryTemplateTagsValue(12)=ParentID
		aryTemplateTagsName( 13)="article/comment/rev"
		aryTemplateTagsValue(13)=""
		Dim objRS
		Set objRS=objConn.Execute("SELECT [comm_ID],[log_ID],[comm_AuthorID],[comm_Author],[comm_Content],[comm_Email],[comm_HomePage],[comm_PostTime],[comm_IP],[comm_Agent],[comm_Reply],[comm_LastReplyIP],[comm_LastReplyTime],[comm_ParentID],[comm_ParentCount],[comm_Meta] FROM [blog_Comment] WHERE [comm_ParentID]=" & ID)
		Set ChildTemplate=New TComment
		Dim intCount
		intCount=0
		Do Until objRS.bof Or objRS.eof
			If ChildTemplate.LoadInfoById(objRS("comm_ID"))=True Then
				intCount=intCount+1
				ChildTemplate.Count=intCount
				aryTemplateTagsValue(13)=aryTemplateTagsValue(13)&ChildTemplate.MakeTemplate(template,True)
			End If
			objRS.movenext
		Loop

		'plugin node
		Call Filter_Plugin_TComment_MakeTemplate_TemplateTags(aryTemplateTagsName,aryTemplateTagsValue)

		j=UBound(aryTemplateTagsName)
		For i=1 to j
			html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
		Next

		MakeTemplate=html

	End Function

	Private Sub Class_Initialize()

		Set Meta=New TMeta

	End Sub

End Class
'*********************************************************




'*********************************************************
' 目的：    定义TTrackBack类
' 输入：    无
' 返回：    无
'*********************************************************
Class TTrackBack

	Public ID
	Public log_ID

	Public URL
	Public Title
	Public Blog
	Public Excerpt

	Public PostTime
	Public IP
	Public Agent

	Public Count
	Public Meta

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property

	Public html

	Public Property Get UrlForAntiSpam
		UrlForAntiSpam=URLEncodeForAntiSpam(Url)
	End Property

	Public Property Get HtmlExcerpt
		HtmlExcerpt=TransferHTML(Excerpt,"[enter]")
	End Property

	Public TbXML

	Private Function ReturnTbXML(strMsg)

		Dim strXML

		strXML="<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><response><error>%e</error><message>%m</message></response>"

		If strMsg="undiscovered" Then'未发现相应ID
			strXML=Replace(strXML,"%e","1")
			strXML=Replace(strXML,"%m",strMsg)
		ElseIf strMsg="repetition" Then'重复PING
			strXML=Replace(strXML,"%e","1")
			strXML=Replace(strXML,"%m",strMsg)
		Elseif strMsg="invalid parameter" Then'非法参数
			strXML=Replace(strXML,"%e","1")
			strXML=Replace(strXML,"%m",strMsg)
		Elseif strMsg="none data" Then'无数据
			strXML=Replace(strXML,"%e","1")
			strXML=Replace(strXML,"%m",strMsg)
		Else'PING 成功
			strXML=Replace(strXML,"%e","0")
			strXML=Replace(strXML,"%m",strMsg)
		End If

		TbXML=strXML

		'Response.ContentType = "text/html"

	End Function


	Public Function Post()

		Call Filter_Plugin_TTrackBack_Post(ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent,MetaString)

		Dim objRS

		'Call ReturnTbXML("undiscovered"):Exit Function
		Call CheckParameter(log_ID,"int",0)

		If IsDate(PostTime)=False Then PostTime=GetTime(Now())
		IP=Request.ServerVariables("REMOTE_ADDR")
		Agent=Request.ServerVariables("HTTP_USER_AGENT")

		IP=FilterSQL(IP)
		Agent=FilterSQL(Agent)

		URL=FilterSQL(URL)
		Title=FilterSQL(Title)
		Blog=FilterSQL(Blog)
		Excerpt=FilterSQL(Excerpt)

		Blog=TransferHTML(Blog,"[html-format]")
		Title=TransferHTML(Title,"[html-format]")
		Excerpt=TransferHTML(Excerpt,"[html-format][nohtml]")
		URL=TransferHTML(URL,"[html-format]")

		'log_ID不能为0
		If (log_ID=0) Then Call ReturnTbXML("invalid parameter"):Post=False:Exit Function
		If Len(URL)=0 Then Call ReturnTbXML("none data"):Post=False:Exit Function
		If Len(URL)>ZC_HOMEPAGE_MAX Then Call ReturnTbXML("url is long"):Post=False:Exit Function

		If Len(Blog)>ZC_EMAIL_MAX Then Call ReturnTbXML("name is long"):Post=False:Exit Function
		If Len(Blog)=0 Then Blog="Unknow"
		If Len(Excerpt)=0 Then Excerpt=""
		If Len(Excerpt)>ZC_TB_EXCERPT_MAX Then Excerpt=Left(Excerpt,ZC_TB_EXCERPT_MAX)&"..."
		If Len(Title)>ZC_HOMEPAGE_MAX Then Call ReturnTbXML("title is long"):Post=False:Exit Function
		If Len(Title)=0 Then Title=URL


	'检查ID是否存在
		'Set objRS=objConn.Execute("SELECT * FROM [blog_Article] WHERE [log_ID]=" & log_ID)
		'If (Not objRS.bof) And (Not objRS.eof) Then
		'Else
		'	objRS.close
		'	Call returnTbXML("undiscovered")
		'	Exit Function
		'End If
		'objRS.Close
		'Set objRS=Nothing

	'检查是否已TB过
		Set objRS=objConn.Execute("SELECT * FROM [blog_TrackBack] WHERE [log_ID]=" & log_ID & " and [tb_url]='" & URL & "'")
		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.close
			Call returnTbXML("repetition")
			Exit Function
		End If
		objRS.Close
		Set objRS=Nothing

	'接收TB

		If ID=0 Then
			objConn.Execute("INSERT INTO [blog_TrackBack]([log_ID],[tb_URL],[tb_Title],[tb_Blog],[tb_Excerpt],[tb_PostTime],[tb_IP],[tb_Agent],[tb_Meta]) VALUES ("&log_ID&",'"&URL&"','"&Title&"','"&Blog&"','"&Excerpt&"','"&PostTime&"','"&IP&"','"&Agent&"','"&MetaString&"')")
		Else
			objConn.Execute("UPDATE [blog_TrackBack] SET [log_ID]="&log_ID&", [tb_URL]='"&URL&"',[tb_Excerpt]='"&Excerpt&"',[tb_Title]='"&Title&"',[tb_Blog]='"&Blog&"',[tb_IP]='"&IP&"',[tb_PostTime]='"&PostTime&"',[tb_Agent]='"&Agent&"',[tb_Meta]='"&MetaString&"' WHERE [tb_ID] =" & ID)
		End If
		Call returnTbXML("succeed")

		Post=True

	End Function

	Public Function Del()

		Call Filter_Plugin_TTrackBack_Del(ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent,MetaString)

		Call CheckParameter(ID,"int",0)
		If (ID=0) Then Del=False:Exit Function
		objConn.Execute("DELETE FROM [blog_TrackBack] WHERE [tb_ID] =" & ID)
		Del=True
	End Function


	Function Send(strAddress)

		Dim strSendTB
		strSendTB = "title=" & Server.URLEncode(Title) & "&url=" & Server.URLEncode(URL) & "&excerpt=" & Server.URLEncode(Excerpt) & "&blog_name=" & Server.URLEncode(Blog)

		Dim objPing
		Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objPing.SetTimeOuts 10000, 10000, 10000, 10000

		objPing.open "POST",strAddress,False

		objPing.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objPing.send strSendTB
		'Response.ContentType = "text/xml"
		'Response.Clear
		'Response.Write objPing.responseXML.xml
		Set objPing = Nothing
		Send=True

	End Function


	Public Function LoadInfoByID(tb_ID)

		Call CheckParameter(tb_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [tb_ID],[log_ID],[tb_URL],[tb_Title],[tb_Blog],[tb_Excerpt],[tb_PostTime],[tb_IP],[tb_Agent],[tb_Meta] FROM [blog_TrackBack] WHERE [tb_ID]=" & tb_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("tb_ID")
			log_ID=objRS("log_ID")
			URL=objRS("tb_URL")
			Title=objRS("tb_Title")
			Blog=objRS("tb_Blog")
			Excerpt=objRS("tb_Excerpt")
			PostTime=objRS("tb_PostTime")
			IP=objRS("tb_IP")
			Agent=objRS("tb_Agent")
			MetaString=objRS("tb_Meta")

			LoadInfoByID=True

		End If

		objRS.Close
		Set objRS=Nothing

		If IsNull(Excerpt) Then Excerpt=""

		Call Filter_Plugin_TTrackBack_LoadInfoByID(ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent,MetaString)

	End Function


	Public Function LoadInfoByArray(aryTbInfo)

		If IsArray(aryTbInfo)=True Then

			ID=aryTbInfo(0)
			log_ID=aryTbInfo(1)
			URL=aryTbInfo(2)
			Title=aryTbInfo(3)
			Blog=aryTbInfo(4)
			Excerpt=aryTbInfo(5)
			PostTime=aryTbInfo(6)
			IP=aryTbInfo(7)
			Agent=aryTbInfo(8)
			MetaString=aryTbInfo(9)

		End If

		If IsNull(Excerpt) Then Excerpt=""

		LoadInfoByArray=True

		Call Filter_Plugin_TTrackBack_LoadInfoByArray(ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent,MetaString)

	End Function


	Public Function MakeTemplate(strT)

		Dim html,i,j
		html=strT

		'plugin node
		Call Filter_Plugin_TTrackBack_MakeTemplate_Template(html)

		Dim aryTemplateTagsName()
		Dim aryTemplateTagsValue()

		ReDim aryTemplateTagsName(7)
		ReDim aryTemplateTagsValue(7)

		aryTemplateTagsName(  1)="article/trackback/id"
		aryTemplateTagsValue( 1)=ID
		aryTemplateTagsName(  2)="article/trackback/name"
		aryTemplateTagsValue( 2)=Blog
		aryTemplateTagsName(  3)="article/trackback/url"
		aryTemplateTagsValue( 3)=UrlForAntiSpam
		aryTemplateTagsName(  4)="article/trackback/title"
		aryTemplateTagsValue( 4)=Title
		aryTemplateTagsName(  5)="article/trackback/posttime"
		aryTemplateTagsValue( 5)=PostTime
		aryTemplateTagsName(  6)="article/trackback/content"
		aryTemplateTagsValue( 6)=HtmlExcerpt
		aryTemplateTagsName(  7)="article/trackback/count"
		aryTemplateTagsValue( 7)=Count

		'plugin node
		Call Filter_Plugin_TTrackBack_MakeTemplate_TemplateTags(aryTemplateTagsName,aryTemplateTagsValue)

		j=UBound(aryTemplateTagsName)
		For i=1 to j
			html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
		Next

		MakeTemplate=html

	End Function

	Private Sub Class_Initialize()

		Set Meta=New TMeta

	End Sub


End Class
'*********************************************************




'*********************************************************
' 目的：    定义TUpLoadFile类
' 输入：    无
' 返回：    无
'*********************************************************
Class TUpLoadFile

	Public ID
	Public AuthorID

	Public FileSize
	Public FileName
	Public PostTime
	Public Stream
	Public DirByTime
	Public FileIntro

	Public Meta

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property

	Public html

	Private FUploadType
	Public Property Let UploadType(strUploadType)
		If (strUploadType="Stream") Then
			FUploadType=strUploadType
		Else
			FUploadType="Form"
		End If
	End Property
	Public Property Get UploadType
		If IsEmpty(FUploadType)=True Then
			UploadType="Form"
		Else
			UploadType = FUploadType
		End If
	End Property

	Public Function LoadInfoByID(ul_ID)

		Call CheckParameter(ul_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [ul_ID],[ul_AuthorID],[ul_FileSize],[ul_FileName],[ul_PostTime],[ul_FileIntro],[ul_DirByTime],[ul_Meta] FROM [blog_UpLoad] WHERE [ul_ID]=" & ul_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("ul_ID")
			AuthorID=objRS("ul_AuthorID")
			FileSize=objRS("ul_FileSize")
			FileName=objRS("ul_FileName")
			PostTime=objRS("ul_PostTime")
			FileIntro=objRS("ul_FileIntro")
			DirByTime=objRS("ul_DirByTime")
			MetaString=objRS("ul_Meta")

			'If IsNull(DirByTime) Or DirByTime="" Then DirByTime=False

			LoadInfobyID=True

		End If

		objRS.Close
		Set objRS=Nothing

		Call Filter_Plugin_TUpLoadFile_LoadInfoByID(ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Meta)

	End Function


	Public Function LoadInfoByArray(aryULInfo)

		If IsArray(aryULInfo)=True Then

			ID=aryULInfo(0)
			AuthorID=aryULInfo(1)
			FileSize=aryULInfo(2)
			FileName=aryULInfo(3)
			PostTime=aryULInfo(4)
			FileIntro=aryULInfo(5)
			DirByTime=aryULInfo(6)
			MetaString=aryULInfo(7)

			'If IsNull(DirByTime) Or DirByTime="" Then DirByTime=False

		End If

		LoadInfoByArray=True

		Call Filter_Plugin_TUpLoadFile_LoadInfoByArray(ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Meta)

	End Function


	Private Function UpLoad_Form()

		Dim i,j
		Dim x,y,z
		Dim intFormSize
		Dim binFormData
		Dim strFileName

		Dim s,t

		intFormSize = Request.TotalBytes
		binFormData = Request.BinaryRead(intFormSize)

		i=0
		i=InstrB(binFormData,ChrB(13)&ChrB(10)&ChrB(13)&ChrB(10))
		If i>0 Then i=i+3
		t=InstrB(binFormData,ChrB(13)&ChrB(10))
		s=Midb(binFormData,1,t)
		t=InstrB(binFormData,ChrB(13)&ChrB(10)&s)
		j=t

		If Len(Request.QueryString("filename"))>0 Then
			strFileName=Request.QueryString("filename")
		Else
			x=InstrB(binFormData,ChrB(&H66)&ChrB(&H69)&ChrB(&H6C)&ChrB(&H65)&ChrB(&H6E)&ChrB(&H61)&ChrB(&H6D)&ChrB(&H65)&ChrB(&H3D)&ChrB(&H22))
			y=InstrB(x+11,binFormData,ChrB(&H22))
			For z=1 to y-x-10
				strFileName=strFileName & Chr(AscB(MidB(binFormData,x+z+9,1)))
			Next
		End If

		Dim objStreamUp
		Set objStreamUp = Server.CreateObject("ADODB.Stream")

		With objStreamUp
			.Type = adTypeBinary
			.Mode = adModeReadWrite
			.Open
			.Position = 0
			.Write binFormData
			.Position = i
			Stream=.Read(j-i-1)
			.Close
		End With

		FileName=strFileName
		FileSize=LenB(Stream)

	End Function


	Private Function UpLoad_Stream()

		FileSize=LenB(Stream)

	End Function


	Public Function UpLoad(bolAutoName)

		Call Filter_Plugin_TUpLoadFile_UpLoad(ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Meta)

		DirByTime=ZC_UPLOAD_DIRBYMONTH

		If UploadType="Form" Then
			Call UpLoad_Form()
		ElseIf UploadType="Stream" Then
			Call UpLoad_Stream()
		End If

		If InStrRev(FileName,"\")>0 Then
			FileName=Mid(FileName,InStrRev(FileName,"\")+1)
		End If

		If InStrRev(FileName,"/")>0 Then
			FileName=Mid(FileName,InStrRev(FileName,"/")+1)
		End If

		FileName=TransferHTML(FileName,"[filename]")

		'超出类型限制
		If Not CheckRegExp(LCase(FileName),"\.("& ZC_UPLOAD_FILETYPE &")$") Then Call ShowError(26)

		'超出大小限制
		If FileSize>ZC_UPLOAD_FILESIZE Then Call ShowError(27)

		FileName=FilterSQL(FileName)
		If bolAutoName=True Then
			Randomize
			FileName=Year(GetTime(Now())) & Right("0"&Month(GetTime(Now())),2) & Right("0"&Day(GetTime(Now())),2) & Right("0"&Hour(GetTime(Now())),2) & Right("0"&Minute(GetTime(Now())),2) & Right("0"&Second(GetTime(Now())),2) & Int(9 * Rnd) & Int(9 * Rnd) & Int(9 * Rnd) & Int(9 * Rnd) & Right(FileName,Len(FileName)-InStrRev(FileName,".")+1)
		End If

		FileIntro=FilterSQL(FileIntro)

		Dim objRS
		Set objRS=objConn.Execute("SELECT * FROM [blog_UpLoad] WHERE [ul_FileName] = '" & FileName & "'")

		If (Not objRS.bof) And (Not objRS.eof) Then
			'不能重名
			 Call ShowError(28)
		Else

			PostTime=GetTime(Now())

			objConn.Execute("INSERT INTO [blog_UpLoad]([ul_AuthorID],[ul_FileSize],[ul_FileName],[ul_PostTime],[ul_FileIntro],[ul_DirByTime],[ul_Meta]) VALUES ("& AuthorID &","& FileSize &",'"& FileName &"','"& PostTime &"','"&FileIntro&"',"&CInt(DirByTime)&",'"&MetaString&"')")

			Dim strUPLOADDIR
			If ZC_UPLOAD_DIRBYMONTH Then
				CreatDirectoryByCustomDirectory(ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now())))
				strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now()))
			Else
				strUPLOADDIR = ZC_UPLOAD_DIRECTORY
			End If

			Dim objStreamFile
			Set objStreamFile = Server.CreateObject("ADODB.Stream")

			objStreamFile.Type = adTypeBinary
			objStreamFile.Mode = adModeReadWrite
			objStreamFile.Open
			objStreamFile.Write Stream

			objStreamFile.SaveToFile BlogPath & "zb_users/"& strUPLOADDIR &"/" & FileName,adSaveCreateOverWrite
			objStreamFile.Close

		End If

		UpLoad=True

	End Function


	Public Function Del()

		Call Filter_Plugin_TUpLoadFile_Del(ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Meta)

		Call CheckParameter(ID,"int",0)

		Dim objRS,strFilePath

		Set objRS=objConn.Execute("SELECT * FROM [blog_UpLoad] WHERE [ul_ID] = " & ID)

		If (Not objRS.bof) And (Not objRS.eof) Then


			Dim fso
			Set fso = CreateObject("Scripting.FileSystemObject")

			strFilePath = BlogPath & "zb_users/"& ZC_UPLOAD_DIRECTORY &"/" & objRS("ul_FileName")

			If fso.FileExists( strFilePath ) Then
				fso.DeleteFile( strFilePath )
			End If

			strFilePath = BlogPath & "zb_users/"& ZC_UPLOAD_DIRECTORY & "/" & Year(objRS("ul_PostTime")) & "/" & Month(objRS("ul_PostTime")) &"/" & objRS("ul_FileName")
			If fso.FileExists( strFilePath ) Then
				fso.DeleteFile( strFilePath )
			End If

			objConn.Execute("DELETE FROM [blog_UpLoad] WHERE [ul_ID] =" & ID)

		Else

			Exit Function

		End If

		objRS.Close
		Set objRS=Nothing

		Del=True

	End Function

	Public Property Get FullUrlPathName

		Dim strUPLOADDIR
		If ZC_UPLOAD_DIRBYMONTH Then
			strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now()))
		Else
			strUPLOADDIR = ZC_UPLOAD_DIRECTORY
		End If

		FullUrlPathName=ZC_BLOG_HOST & strUPLOADDIR & "/" & FileName
	End Property

	Private Sub Class_Initialize()

		Set Meta=New TMeta

	End Sub


End Class
'*********************************************************




'*********************************************************
' 目的：    定义TTag类
' 输入：    无
' 返回：    无
'*********************************************************
Class TTag

	Public ID
	Public Name
	Public Intro
	Public Order
	Public Count
	Public Alias
	Public ParentID

	Public Meta
	Public TemplateName

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property

	Public html

	Public Property Get EncodeName
		EncodeName = Server.URLEncode(Name)
	End Property

	Public Property Get Url

		'plugin node
		bAction_Plugin_TTag_Url=False
		For Each sAction_Plugin_TTag_Url in Action_Plugin_TTag_Url
			If Not IsEmpty(sAction_Plugin_TTag_Url) Then Call Execute(sAction_Plugin_TTag_Url)
			If bAction_Plugin_TTag_Url=True Then Exit Property
		Next

		Url = ZC_BLOG_HOST & "catalog.asp?"& "tags=" & Server.URLEncode(Name)

		'If CheckPluginState("STACentre") Then
		'	'Dim objTag : Set objTag=New STACentre_Tags : If objTag.LoadInfoById(ID) Then Url=objTag.Url : Set objTag=Nothing
		'End If

		Call Filter_Plugin_TTag_Url(Url)
	End Property

	Public Property Get HtmlUrl
		HtmlUrl=TransferHTML(Url,"[html-format]")
	End Property

	Public Property Get HtmlIntro
		HtmlIntro=TransferHTML(Intro,"[html-format]")
	End Property

	Public Property Get HtmlName
		HtmlName=TransferHTML(Name,"[html-format]")
	End Property

	Public Property Get RssUrl
		RssUrl = ZC_BLOG_HOST & "sydication.asp?tags=" & ID
	End Property

	Public Function Post()

		Call Filter_Plugin_TTag_Post(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,MetaString)

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(Order,"int",0)
		Call CheckParameter(ParentID,"int",0)

		Name=FilterSQL(Name)
		Name=TransferHTML(Name,"[normalname]")
		If Len(Name)=0 Then Post=False:Exit Function

		Intro=FilterSQL(Intro)
		Alias=FilterSQL(Alias)
		'Intro=TransferHTML(Intro,"[html-format]")

		TemplateName=UCase(FilterSQL(TemplateName))
		If TemplateName="CATALOG" Then TemplateName=""

		If ID=0 Then
			objConn.Execute("INSERT INTO [blog_Tag]([tag_Name],[tag_Order],[tag_Intro],[tag_ParentID],[tag_URL],[tag_Template],[tag_Meta]) VALUES ('"&Name&"',"&Order&",'"&Intro&"',"&ParentID&",'"&Alias&"','"&TemplateName&"','"&MetaString&"')")
			Dim objRS
			Set objRS=objConn.Execute("SELECT MAX([tag_ID]) FROM [blog_Tag]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
			End If
			Set objRS=Nothing
		Else
			objConn.Execute("UPDATE [blog_Tag] SET [tag_Name]='"&Name&"',[tag_Order]="&Order&",[tag_Intro]='"&Intro&"',[tag_ParentID]="&ParentID&",[tag_URL]='"&Alias&"',[tag_Template]='"&TemplateName&"',[tag_Meta]='"&MetaString&"' WHERE [tag_ID] =" & ID)
		End If

		Post=True

	End Function


	Public Function LoadInfoByID(tag_ID)

		Call CheckParameter(tag_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [tag_ID],[tag_Name],[tag_Intro],[tag_Order],[tag_Count],[tag_ParentID],[tag_URL],[tag_Template],[tag_Meta] FROM [blog_Tag] WHERE [tag_ID]=" & tag_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("tag_ID")
			Name=objRS("tag_Name")
			Intro=objRS("tag_Intro")
			Order=objRS("tag_Order")
			Count=objRS("tag_Count")
			ParentID=objRS("tag_ParentID")
			Alias=objRS("tag_URL")
			TemplateName=objRS("tag_Template")
			MetaString=objRS("tag_Meta")

			LoadInfoByID=True

		End If

		objRS.Close
		Set objRS=Nothing

		If IsNull(Intro) Then Intro=""

		Call Filter_Plugin_TTag_LoadInfoByID(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,MetaString)

	End Function

	Public Function LoadInfoByArray(aryTagInfo)

		If IsArray(aryTagInfo)=True Then
			ID=aryTagInfo(0)
			Name=aryTagInfo(1)
			Intro=aryTagInfo(2)
			Order=aryTagInfo(3)
			Count=aryTagInfo(4)
			ParentID=aryTagInfo(5)
			Alias=aryTagInfo(6)
			TemplateName=aryTagInfo(7)
			MetaString=aryTagInfo(8)
		End If

		If IsNull(Intro) Then Intro=""
		If IsNull(Alias) Then Alias=""

		LoadInfoByArray=True


		Call Filter_Plugin_TTag_LoadInfoByArray(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,MetaString)

	End Function


	Public Function Del()

		Call Filter_Plugin_TTag_Del(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,MetaString)

		Call CheckParameter(ID,"int",0)
		If (ID=0) Then Del=False:Exit Function

		Dim s
		Dim i
		Dim objRS

		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn
		objRS.Source=""

		objRS.Open("SELECT [log_ID],[log_tag] FROM [blog_Article] WHERE [log_Tag] LIKE '%{" & ID & "}%'")

		If (Not objRS.bof) And (Not objRS.eof) Then
			Do While Not objRS.eof
				i=objRS("log_ID")
				s=objRS("log_tag")
				s=Replace(s,"{"& ID &"}","")
				objConn.Execute("UPDATE [blog_Article] SET [log_tag]='"& s &"' WHERE [log_ID] =" & i)
				objRS.MoveNext
			Loop
		End If
		objRS.Close

		objConn.Execute("DELETE FROM [blog_Tag] WHERE [tag_ID] =" & ID)
		Del=True
	End Function


	Public Function MakeTemplate(s)

		Dim html,i,j
		html=s

		'plugin node
		Call Filter_Plugin_TTag_MakeTemplate_Template(html)

		Dim aryTemplateTagsName()
		Dim aryTemplateTagsValue()

		ReDim aryTemplateTagsName(7)
		ReDim aryTemplateTagsValue(7)

		aryTemplateTagsName(  1)="article/tag/id"
		aryTemplateTagsValue( 1)=ID
		aryTemplateTagsName(  2)="article/tag/name"
		aryTemplateTagsValue( 2)=HtmlName
		aryTemplateTagsName(  3)="article/tag/intro"
		aryTemplateTagsValue( 3)=HtmlIntro
		aryTemplateTagsName(  4)="article/tag/count"
		aryTemplateTagsValue( 4)=Count
		aryTemplateTagsName(  5)="article/tag/url"
		aryTemplateTagsValue( 5)=HtmlUrl
		aryTemplateTagsName(  6)="article/tag/encodename"
		aryTemplateTagsValue( 6)=EncodeName

		'plugin node
		Call Filter_Plugin_TTag_MakeTemplate_TemplateTags(aryTemplateTagsName,aryTemplateTagsValue)

		j=UBound(aryTemplateTagsName)
		For i=1 to j
			If IsNull(aryTemplateTagsValue(i))=False Then
				html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
			Else
				html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>","")
			End If
		Next

		MakeTemplate=html

	End Function

	Private Sub Class_Initialize()

		Set Meta=New TMeta

	End Sub


End Class
'*********************************************************




'*********************************************************
' 目的：    定义TKeyWord类
' 输入：    无
' 返回：    无
'*********************************************************
Class TKeyWord

	Public ID
	Public Name
	Public Intro
	Public Url

	Public Function Post()

		Call CheckParameter(ID,"int",0)

		Name=FilterSQL(Name)
		Name=TransferHTML(Name,"[normalname]")

		If Len(Name)=0 Then Post=False:Exit Function

		Intro=FilterSQL(Intro)
		Intro=TransferHTML(Intro,"[html-format]")

		Url=FilterSQL(Url)
		If Len(Url)=0 Then Post=False:Exit Function
		If Not CheckRegExp(Url,"[homepage]") Then Call  ShowError(30)

		If ID=0 Then
			objConn.Execute("INSERT INTO [blog_Keyword]([key_Name],[key_URL],[key_Intro]) VALUES ('"&Name&"','"&Url&"','"&Intro&"')")
		Else
			objConn.Execute("UPDATE [blog_Keyword] SET [key_Name]='"&Name&"',[key_URL]='"&Url&"',[key_Intro]='"&Intro&"' WHERE [key_ID] =" & ID)
		End If

		Post=True

	End Function


	Public Function LoadInfoByID(key_ID)

		Call CheckParameter(key_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [key_ID],[key_Name],[key_Intro],[key_Url] FROM [blog_Keyword] WHERE [key_ID]=" & key_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("key_ID")
			Name=objRS("key_Name")
			Intro=objRS("key_Intro")
			Url=objRS("key_Url")

			LoadInfoByID=True

		End If

		objRS.Close
		Set objRS=Nothing

		If IsNull(Intro) Then Intro=""

	End Function


	Public Function LoadInfoByArray(aryKeyWordInfo)

		If IsArray(aryKeywordInfo)=True Then

			ID=aryKeyWordInfo(0)
			Name=aryKeyWordInfo(1)
			Intro=aryKeyWordInfo(2)
			Url=aryKeyWordInfo(3)

		End If

		If IsNull(Intro) Then Intro=""

		LoadInfoByArray=True

	End Function


	Public Function Del()

		Call CheckParameter(ID,"int",0)
		If (ID=0) Then Del=False:Exit Function

		objConn.Execute("DELETE FROM [blog_Keyword] WHERE [key_ID] =" & ID)
		Del=True

	End Function


End Class
'*********************************************************




'*********************************************************
' 目的：    定义TRss2Export类 New版
' 输入：    无
' 返回：    无
'*********************************************************
Class TNewRss2Export

	Public TimeZone

	Public Property Get xml
		xml = objXMLdoc.xml
	End Property

	public FstrWebLink
	public FstrAuthor

	Public Property Get WebLink
		WebLink = FstrWebLink
	End Property

	Public Property Let WebLink(strWebLink)
		FstrWebLink = strWebLink
	End Property

	Public Property Get Author
		Author = FstrAuthor
	End Property

	Public Property Let Author(strAuthor)
		FstrAuthor = strAuthor
	End Property

	Private objXMLdoc

	Private objXMLrss

	Private objXMLchannel


	Public Function AddChannelAttribute(title,value)

		Dim objXMLitem
		Set objXMLitem = objXMLdoc.createElement(title)

		If title="pubDate" Then value=ParseDateForRFC822(value)

		objXMLitem.text=value
		objXMLchannel.AppendChild(objXMLitem)

		AddChannelAttribute=True

	End Function


	Public Function AddItem(title,author,link,pubDate,guid,description,category,comments,wfw_comment,wfw_commentRss,trackback_ping)

		Dim objXMLitem
		Set objXMLitem = objXMLdoc.createElement("item")
		Dim objXMLcdata

		If(Len(title)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("title"))
			objXMLitem.selectSingleNode("title").text=title
		End If
		If(Len(author)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("author"))
			objXMLitem.selectSingleNode("author").text=author
		End If
		If(Len(link)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("link"))
			objXMLitem.selectSingleNode("link").text=link
		End If
		If(Len(pubDate)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("pubDate"))
			objXMLitem.selectSingleNode("pubDate").text=ParseDateForRFC822(pubDate)
		End If
		If(Len(guid)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("guid"))
			objXMLitem.selectSingleNode("guid").text=guid
		End If
		If(Len(description)>0) Then

			objXMLitem.AppendChild(objXMLdoc.createElement("description"))
			Set objXMLcdata = objXMLdoc.createNode("cdatasection", "","")
			objXMLcdata.NodeValue=description
			objXMLitem.selectSingleNode("description").AppendChild(objXMLcdata)

			Set objXMLcdata = Nothing

		End If
		If(Len(category)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("category"))
			objXMLitem.selectSingleNode("category").text=category
		End If

		If(Len(comments)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("comments"))
			objXMLitem.selectSingleNode("comments").text=comments
		End If
		If(Len(wfw_comment)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("wfw:comment"))
			objXMLitem.selectSingleNode("wfw:comment").text=wfw_comment
		End If
		If(Len(wfw_commentRss)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("wfw:commentRss"))
			objXMLitem.selectSingleNode("wfw:commentRss").text=wfw_commentRss
		End If
		If(Len(trackback_ping)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("trackback:ping"))
			objXMLitem.selectSingleNode("trackback:ping").text=trackback_ping
		End If

		objXMLchannel.AppendChild(objXMLitem)

		AddItem=True

	End Function

	Public Function Execute()

		'Response.ContentType = "text/html"
		Response.ContentType = "text/xml"
		Response.Clear
		Response.Write xml

		Execute=True

	End Function

	Public Function SaveToFile(strFileName)

		objXMLdoc.save(strFileName)

		SaveToFile=True

	End Function


	Function ParseDateForRFC822(dtmDate)

		Dim dtmDay, dtmWeekDay, dtmMonth, dtmYear
		Dim dtmHours, dtmMinutes, dtmSeconds

		Select Case WeekDay(dtmDate)
			Case 1:dtmWeekDay="Sun"
			Case 2:dtmWeekDay="Mon"
			Case 3:dtmWeekDay="Tue"
			Case 4:dtmWeekDay="Wed"
			Case 5:dtmWeekDay="Thu"
			Case 6:dtmWeekDay="Fri"
			Case 7:dtmWeekDay="Sat"
		End Select

		Select Case Month(dtmDate)
			Case 1:dtmMonth="Jan"
			Case 2:dtmMonth="Feb"
			Case 3:dtmMonth="Mar"
			Case 4:dtmMonth="Apr"
			Case 5:dtmMonth="May"
			Case 6:dtmMonth="Jun"
			Case 7:dtmMonth="Jul"
			Case 8:dtmMonth="Aug"
			Case 9:dtmMonth="Sep"
			Case 10:dtmMonth="Oct"
			Case 11:dtmMonth="Nov"
			Case 12:dtmMonth="Dec"
		End Select

		dtmYear = Year(dtmDate)
		dtmDay = Right("00" & Day(dtmDate),2)

		dtmHours = Right("00" & Hour(dtmDate),2)
		dtmMinutes = Right("00" & Minute(dtmDate),2)
		dtmSeconds = Right("00" & Second(dtmDate),2)

		ParseDateForRFC822 = dtmWeekDay & ", " & dtmDay &" " & dtmMonth & " " & dtmYear & " " & dtmHours & ":" & dtmMinutes & ":" & dtmSeconds & " " & TimeZone

	End Function

	' 类初始化
	Private Sub Class_Initialize()

		On Error Resume Next

		'对objXMLdoc进行初始化，如不能建对象则报错
		Set objXMLdoc =Server.CreateObject("Microsoft.XMLDOM")

		If Err.Number<>0 Then

		End If

		Dim objPI

		'Set objPI = objXMLdoc.createProcessingInstruction("xml-stylesheet","type=""text/css"" href=""zb_system/css/rss.css""")
		'objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
		'Set objPI = Nothing

		Set objPI = objXMLdoc.createProcessingInstruction("xml-stylesheet","type=""text/xsl"" href=""zb_system/css/rss.xslt""")
		objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
		Set objPI = Nothing

		Set objPI = objXMLdoc.createProcessingInstruction("xml","version=""1.0"" encoding=""UTF-8"" standalone=""yes""")
		objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
		Set objPI = Nothing

		Set objXMLrss = objXMLdoc.createElement("rss")

		Set objXMLchannel = objXMLdoc.createElement("channel")


		objXMLrss.AppendChild(objXMLchannel)
		objXMLdoc.AppendChild(objXMLrss)

		objXMLrss.setAttribute "version","2.0"
		objXMLrss.setAttribute "xmlns:dc","http://purl.org/dc/elements/1.1/"
		objXMLrss.setAttribute "xmlns:trackback","http://madskills.com/public/xml/rss/module/trackback/"
		objXMLrss.setAttribute "xmlns:wfw","http://wellformedweb.org/CommentAPI/"
		objXMLrss.setAttribute "xmlns:slash","http://purl.org/rss/1.0/modules/slash/"


	End Sub

	' 类释放
	Private Sub Class_Terminate()

		Set objXMLrss = Nothing
		Set objXMLdoc  = Nothing

	End Sub

End Class
'*********************************************************




'*********************************************************
' 目的：    定义TMeta
' 输入：    无
' 返回：    无
'*********************************************************
Class TMeta

'@猪: 关于分隔用的字符串请重新考虑下:
'1. {name|a:123|b:345} by 月木
'2. 在存储时过滤掉可能引起 Split 出错的字符串.

	Dim meta_split_string_1
	Dim meta_split_string_2

	Dim Names()
	Dim Values()

	Public Property Get Count
		Count = UBound(names)+1
	End Property

	Public Function Save()

		Dim n,v
		If UBound(names)>-1 Then

			'n=Join(Names,"</name><name>")
			'n="<name>"&n&"</name>"
			'v=Join(Values,"</value><value>")
			'v="<value>"&v&"</value>"
			'Save="<names>"&n&"</names>" & "<values>"&v&"</values>"
			'Save="<xml>" & Save & "</xml>"

			n=Join(Names,meta_split_string_1)
			v=Join(Values,meta_split_string_1)

			Save=Join(Array(n,v),meta_split_string_2)

		End If

	End Function


	Public Property Get SaveString
		SaveString = Save()
	End Property

	Public Property Let LoadString(s)
		Call Load(s)
	End Property

	Public Function Load(s)

		If IsNull(s)=True Then Exit Function
		If IsEmpty(s)=True Then Exit Function

		Dim x,n,v
		x=Split(s,meta_split_string_2)

		If UBound(x)<>1 Then Exit Function

		n=Split(x(0),meta_split_string_1)
		v=Split(x(1),meta_split_string_1)

		Dim i
		ReDim Names(UBound(n))
		ReDim Values(UBound(n))
		For i=0 To UBound(n)
			Names(i)=n(i)
			Values(i)=v(i)
		Next

	End Function


	Public Function LoadByArray(n,v)

		Dim i
		ReDim Names(UBound(n))
		ReDim Values(UBound(n))
		For i=0 To UBound(n)
			Names(i)=n(i)
			Values(i)=v(i)
		Next

	End Function


	Public Function SetValue(name,value)

		name=LCase(name)

		Dim n,i
		i=0
		For Each n In names
			If n=name Then
				values(i)=value
				Exit function
			End If
			i=i+1
		Next

		i=UBound(names)

		ReDim Preserve Names(i+1)
		ReDim Preserve Values(i+1)

		Names(i+1)=LCase(name)
		Values(i+1)=value

	End Function

	Public Function GetValue(name)

		Dim n,i
		i=0
		For Each n In names
			If n=name Then
				GetValue = values(i)
				Exit function
			End If
			i=i+1
		Next

		GetValue = null
	End Function

	Public Function Remove(name)

		name=LCase(name)

		Dim n()
		Dim v()

		ReDim n(UBound(names))
		ReDim v(UBound(names))

		Dim i,j
		j=0
		For i=0 To UBound(names)
			If names(i)=name Then

			Else
				n(j)=names(i)
				v(j)=values(i)
				j=j+1
			End If
		Next

		ReDim names(j-1)
		ReDim values(j-1)

		For i=0 To j-1
			names(i)=n(i)
			values(i)=v(i)
		Next

	End Function

	Private Sub Class_Initialize()
		ReDim Names(0)
		ReDim Values(0)

		ReDim Names(-1)
		ReDim Values(-1)

		meta_split_string_1=Chr(12) & Chr(9) & Chr(10) & Chr(13) & Chr(10) & Chr(12)
		meta_split_string_2=Chr(12) & Chr(13)&  Chr(10) & Chr(9) & Chr(10) & Chr(12)
	End Sub

End Class
'*********************************************************
%>