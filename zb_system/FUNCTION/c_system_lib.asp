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
	Public FullUrl
	Public Meta
	Public TemplateName
	Public LogTemplate

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property


	Private Ftemplate
	Public Property Let Template(strFileName)
		Ftemplate=GetTemplate("TEMPLATE_" & strFileName)
	End Property
	Public Property Get Template
		If Ftemplate<>"" Then
			Template = Ftemplate
			Exit Property
		Else
			If TemplateName<>"" Then
				Dim s
				s=GetTemplate("TEMPLATE_" &TemplateName)
				If s<>"" Then
					Ftemplate = s
				Else
					Ftemplate=GetTemplate("TEMPLATE_CATALOG")
				End If
			Else
				Ftemplate=GetTemplate("TEMPLATE_CATALOG")
			End If
			Template = Ftemplate
		End If
	End Property

	Public Function GetDefaultTemplateName
		If TemplateName<>"" Then
			GetDefaultTemplateName=TemplateName
		Else
			GetDefaultTemplateName="CATALOG"
		End If
	End Function

	Public Function GetDefaultLogTemplateName
		If LogTemplate<>""  Then
			GetDefaultLogTemplateName=LogTemplate
		Else
			GetDefaultLogTemplateName="SINGLE"
		End IF
	End Function

	Private Ffullregex
	Public Property Let FullRegex(s)
		Ffullregex=s
	End Property
	Public Property Get FullRegex
		If Ffullregex<>"" Then
			FullRegex=Ffullregex
		Else
			FullRegex=ZC_CATEGORY_REGEX
		End If
	End Property

	Public html

	Public Property Get FullPath
		FullPath=ParseCustomDirectoryForPath(FullRegex,ZC_STATIC_DIRECTORY,"","","","","",ID,Name,StaticName)
	End Property

	Public Property Get Url

		'plugin node
		bAction_Plugin_TCategory_Url=False
		For Each sAction_Plugin_TCategory_Url in Action_Plugin_TCategory_Url
			If Not IsEmpty(sAction_Plugin_TCategory_Url) Then Call Execute(sAction_Plugin_TCategory_Url)
			If bAction_Plugin_TCategory_Url=True Then Exit Property
		Next

		Url =ParseCustomDirectoryForUrl(FullRegex,ZC_STATIC_DIRECTORY,"","","","","",ID,Name,StaticName)
		If Right(Url,12)="default.html" Then Url=Left(Url,Len(Url)-12)

		Url=Replace(Replace(Url,"//","/"),":/","://",1,1)

		Call Filter_Plugin_TCategory_Url(Url)

	End Property

	Public Property Get RssUrl
		RssUrl = BlogHost & "feed.asp?cate=" & ID
	End Property

	Public Property Get HtmlName
		HtmlName=TransferHTML(Name,"[html-format]")
	End Property

	Public Property Get HtmlUrl
		HtmlUrl=TransferHTML(Url,"[html-format]")
	End Property


	Public Property Get StaticName
		If IsNull(Alias) Or IsEmpty(Alias) Or Alias="" Then
			StaticName = Name
		Else
			StaticName = Alias
		End If
	End Property

	Public ReCount
	Public ExID '原ID

	Public Function Post()

		Call Filter_Plugin_TCategory_Post(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,LogTemplate,FullUrl,MetaString)

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(Order,"int",0)
		Call CheckParameter(ParentID,"int",0)

		'ID可以为0
		Name=FilterSQL(Name)
		Alias=TransferHTML(Alias,"[directory&file]")
		If Left(Alias,1)="/" Then Alias=Right(Alias,Len(Alias)-1)
		If Right(Alias,1)="/" Then Alias=Left(Alias,Len(Alias)-1)
		Alias=FilterSQL(Alias)
		Intro=FilterSQL(Intro)

		TemplateName=UCase(FilterSQL(TemplateName))
		If TemplateName="CATALOG" Then TemplateName=""

		LogTemplate=UCase(FilterSQL(LogTemplate))
		If LogTemplate="SINGLE" Then LogTemplate=""

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

			objConn.Execute("INSERT INTO [blog_Category]([cate_Name],[cate_Order],[cate_Intro],[cate_ParentID],[cate_Url],[cate_Template],[cate_LogTemplate],[cate_FullUrl],[cate_Meta]) VALUES ('"&Name&"',"&Order&",'"&Intro&"',"&ParentID&",'"&Alias&"','"&TemplateName&"','"&LogTemplate&"','"&FullUrl&"','"&MetaString&"')")

			Dim objRS
			Set objRS=objConn.Execute("SELECT MAX([cate_ID]) FROM [blog_Category]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
			End If
			Set objRS=Nothing

			If ParentID=ID Then
				ParentID=0
				objConn.Execute("UPDATE [blog_Category] set [cate_Name]='"&Name&"',[cate_Order]="&Order&",[cate_Intro]='"&Intro&"',[cate_ParentID]="&ParentID&",[cate_Url]='"&Alias&"',[cate_Template]='"&TemplateName&"',[cate_LogTemplate]='"&LogTemplate&"',[cate_FullUrl]='"&FullUrl&"',[cate_Meta]='"&MetaString&"' WHERE [cate_ID] =" & ID)
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

			objConn.Execute("UPDATE [blog_Category] set [cate_Name]='"&Name&"',[cate_Order]="&Order&",[cate_Intro]='"&Intro&"',[cate_ParentID]="&ParentID&",[cate_Url]='"&Alias&"',[cate_Template]='"&TemplateName&"',[cate_LogTemplate]='"&LogTemplate&"',[cate_FullUrl]='"&FullUrl&"',[cate_Meta]='"&MetaString&"' WHERE [cate_ID] =" & ID)

		End If

		Post=True

	End Function


	Public Function LoadInfoByID(cate_ID)

		Call CheckParameter(cate_ID,"int",0)

		If cate_ID=0 Then
			If BlogConfig.Exists("ZC_UNCATEGORIZED_NAME")=True Then Name=BlogConfig.Read("ZC_UNCATEGORIZED_NAME")
			If BlogConfig.Exists("ZC_UNCATEGORIZED_ALIAS")=True Then Alias=BlogConfig.Read("ZC_UNCATEGORIZED_ALIAS")
			If BlogConfig.Exists("ZC_UNCATEGORIZED_COUNT")=True Then
				Dim intUnCateCount
				intUnCateCount=BlogConfig.Read("ZC_UNCATEGORIZED_COUNT")
				Call CheckParameter(intUnCateCount,"int",0)
				Count=intUnCateCount
			End if
			LoadInfoByID=True
			Exit Function
		End If

		Dim objRS
		Set objRS=objConn.Execute("SELECT [cate_ID],[cate_Name],[cate_Intro],[cate_Order],[cate_Count],[cate_ParentID],[cate_Url],[cate_Template],[cate_LogTemplate],[cate_FullUrl],[cate_Meta] FROM [blog_Category] WHERE [cate_ID]=" & cate_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("cate_ID")
			Name=objRS("cate_Name")
			Alias=objRS("cate_Url")
			Order=objRS("cate_Order")
			Count=objRS("cate_Count")
			ParentID=objRS("cate_ParentID")
			Intro=objRS("cate_Intro")
			TemplateName=objRS("cate_Template")
			LogTemplate=objRS("cate_LogTemplate")
			FullUrl=objRS("cate_FullUrl")
			MetaString=objRS("cate_Meta")
			LoadInfoByID=True

		End If

		objRS.Close
		Set objRS=Nothing

		Call Filter_Plugin_TCategory_LoadInfoByID(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,LogTemplate,FullUrl,MetaString)

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
			LogTemplate=aryCateInfo(8)
			FullUrl=aryCateInfo(9)
			MetaString=aryCateInfo(10)
		End If

		LoadInfoByArray=True

		Call Filter_Plugin_TCategory_LoadInfoByArray(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,LogTemplate,FullUrl,MetaString)

	End Function


	Public Function Del()

		Call Filter_Plugin_TCategory_Del(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,LogTemplate,FullUrl,MetaString)

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
		Name=ZC_MSG059
		ReCount=0
		Order=0
		ExID=-1 '表示无原ID，文章分类无变化
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
	Public FType
	Public Meta
	Public TemplateName

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property

	Private Template_Article_Trackback
	Private Template_Article_Comment
	Private Template_Article_Comment_Pagebar
	Private Template_Article_Commentpost
	Private Template_Article_Tag
	Private Template_Article_Navbar_L
	Private Template_Article_Navbar_R
	Private Template_Article_Commentpost_Verify
	Private Template_Article_Mutuality
	Private Template_Calendar


	Private Disable_Export_Tag
	Private Disable_Export_CMTandTB
	Private Disable_Export_CommentPost
	Private Disable_Export_Mutuality
	Private Disable_Export_NavBar

	Private HasTag
	Private HasCMTandTB
	Private HasMutuality

	Public html
	Public subhtml
	Public subhtml_TemplateName

	Public IsDynamicLoadSildbar
	Public SearchText
	Public CommentsPage

	Public Property Get IsPage
		If FType=1 Then
			IsPage=True
		Else
			IsPage=False
		End If
	End Property

	Private Ftemplate
	Public Property Let Template(strFileName)
		Ftemplate=GetTemplate("TEMPLATE_" & strFileName)
	End Property
	Public Property Get Template
		If Ftemplate<>"" Then
			Template = Ftemplate
			Exit Property
		Else
			If TemplateName<>"" Then
				Dim s
				s=GetTemplate("TEMPLATE_" &TemplateName)
				If s<>"" Then
					Ftemplate = s
					Template = Ftemplate
					Exit Property
				End If
			End If
			If IsPage=False And Categorys(CateID).LogTemplate<>"" Then
				Dim t
				t=GetTemplate("TEMPLATE_" &Categorys(CateID).LogTemplate)
				If t<>"" Then
					Ftemplate = t
					Template = Ftemplate
					Exit Property
				End If
			End If
			If IsPage=True Then
				Ftemplate=GetTemplate("TEMPLATE_PAGE")
			Else
				Ftemplate=GetTemplate("TEMPLATE_SINGLE")
			End If

			Template = Ftemplate
		End If
	End Property

	Public Function GetDefaultTemplateName
		If TemplateName<>"" Then
			GetDefaultTemplateName=TemplateName
		Else
			If IsPage=True Then
				GetDefaultTemplateName="PAGE"
			Else
				If Categorys(CateID).LogTemplate<>"" Then
					GetDefaultTemplateName=Categorys(CateID).LogTemplate
				Else
					GetDefaultTemplateName="SINGLE"
				End If
			End If
		End If
	End Function

	Private Ffullregex
	Public Property Let FullRegex(s)
		Ffullregex=s
	End Property
	Public Property Get FullRegex
		If Ffullregex<>"" Then
			FullRegex=Ffullregex
		Else
			If IsPage=True Then
				If Level>2 Then
					FullRegex=ZC_PAGE_REGEX
				ElseIf Level>1 Then
					FullRegex=ZC_PAGE_AND_ARTICLE_PRIVATE_REGEX
				Else
					FullRegex=ZC_PAGE_AND_ARTICLE_DRAFT_REGEX
				End If
			Else
				If Level>2 Then
					FullRegex=ZC_ARTICLE_REGEX
				ElseIf Level>1 Then
					FullRegex=ZC_PAGE_AND_ARTICLE_PRIVATE_REGEX
				Else
					FullRegex=ZC_PAGE_AND_ARTICLE_DRAFT_REGEX
				End If
			End If
		End If
	End Property


	Public Property Get FullPath
		Call GetUsersbyUserIDList(AuthorID)
		FullPath=ParseCustomDirectoryForPath(FullRegex,ZC_STATIC_DIRECTORY,Categorys(CateID).StaticName,Users(AuthorID).StaticName,Year(PostTime),Month(PostTime),Day(PostTime),ID,StaticName,StaticName)
	End Property


	Private FUrl
	Public Property Get Url

		If FUrl<>"" Then
			Url=FUrl
		Else

			'plugin node
			bAction_Plugin_TArticle_Url=False
			For Each sAction_Plugin_TArticle_Url in Action_Plugin_TArticle_Url
				If Not IsEmpty(sAction_Plugin_TArticle_Url) Then Call Execute(sAction_Plugin_TArticle_Url)
				If bAction_Plugin_TArticle_Url=True Then Exit Property
			Next

			Call GetUsersbyUserIDList(AuthorID)
			FUrl =ParseCustomDirectoryForUrl(FullRegex,ZC_STATIC_DIRECTORY,Categorys(CateID).StaticName,Users(AuthorID).StaticName,Year(PostTime),Month(PostTime),Day(PostTime),ID,StaticName,StaticName)
			If Right(FUrl,12)="default.html" Then FUrl=Left(FUrl,Len(FUrl)-12)
			'If Right(Url,10)="index.html" Then Url=Left(Url,Len(Url)-10)

			FUrl=Replace(Replace(FUrl,"//","/"),":/","://",1,1)

			Call Filter_Plugin_TArticle_Url(FUrl)
			Url=FUrl

		End If

	End Property

	Public Property Get StaticName
		If IsNull(Alias) Or IsEmpty(Alias) Or Alias="" Then
			StaticName = ID
		Else
			StaticName = Alias
		End If
	End Property

	Private FTrackBackKey
	Public Property Get TrackBackKey
		If IsNull(FTrackBackKey) Or IsEmpty(FTrackBackKey) Or FTrackBackKey="" Then
			'FTrackBackKey=Left(MD5(ZC_BLOG_CLSID & CStr(ID) & CStr(TrackBackNums)),8)
			FTrackBackKey=Empty
		End If
		TrackBackKey=FTrackBackKey
	End Property

	Private FCommentKey
	Public Property Get CommentKey
		If IsNull(FCommentKey) Or IsEmpty(FCommentKey) Or FCommentKey="" Then
			FCommentKey=Left(MD5(ZC_BLOG_CLSID & CStr(ID)),8)
		End If
		CommentKey=FCommentKey
	End Property

	Public Property Get TrackBack
		TrackBack = BlogHost & "zb_system/cmd.asp?act=tb&id="& ID &"&key=" & TrackBackKey
	End Property

	Public Property Get PreTrackBack
		PreTrackBack = BlogHost & "zb_system/cmd.asp?act=gettburl&id=" & ID
	End Property

	Public Property Get TrackBackUrl
		TrackBackUrl = TrackBack
	End Property

	Public Property Get CommentUrl
		CommentUrl = Url & "#comment"
	End Property

	Public Property Get WfwComment
		WfwComment = BlogHost
	End Property

	Public Property Get WfwCommentRss
		WfwCommentRss = BlogHost & "feed.asp?cmt=" & ID
	End Property

	Public Property Get WAPUrl
		WAPUrl = BlogHost & "?mod=wap&act=View&id=" & ID
	End Property

	Public Property Get HtmlWAPUrl
		HtmlWAPUrl=TransferHTML(WAPUrl,"[html-format]")
	End Property

	Public Property Get PadUrl
		PadUrl = BlogHost & "?mod=pad&act=View&id=" & ID
	End Property

	Public Property Get HtmlPadUrl
		HtmlPadUrl=TransferHTML(PadUrl,"[html-format]")
	End Property

	Public Property Get CommentPostUrl
		CommentPostUrl = BlogHost & "zb_system/cmd.asp?act=cmt&key=" & CommentKey
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
		HtmlUrl=TransferHTML(TransferHTML(Url,"[zc_blog_host]"),"[html-format]")
	End Property

	Public Property Get TagToName

		Dim t,i,s

		If Tag<>"" Then
			Call GetTagsbyTagIDList(Tag)
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
		Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE [log_ID]=" & log_ID)

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
			FType=objRS("log_Type")
			MetaString=objRS("log_Meta")

			Content=TransferHTML(Content,"[upload][zc_blog_host]")
			Intro=TransferHTML(Intro,"[upload][zc_blog_host]")

			If ZC_POST_STATIC_MODE<>"STATIC" Then FullUrl=Replace(Url,BlogHost,"<#ZC_BLOG_HOST#>")

			PostTime = Year(PostTime) & "-" & Month(PostTime) & "-" & Day(PostTime) & " " & Hour(PostTime) & ":" & Minute(PostTime) & ":" & Second(PostTime)

		Else
			Exit Function
		End If

		objRS.Close
		Set objRS=Nothing

		FUrl=""

		LoadInfobyID=True

		Call Filter_Plugin_TArticle_LoadInfobyID(ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,FullUrl,FType,MetaString)

	End Function



	Public Function LoadInfoByArray(aryArticleInfo)

		'[log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta]

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
			FType=aryArticleInfo(16)
			MetaString=aryArticleInfo(17)

			Content=TransferHTML(Content,"[upload][zc_blog_host]")
			Intro=TransferHTML(Intro,"[upload][zc_blog_host]")

			PostTime = Year(PostTime) & "-" & Month(PostTime) & "-" & Day(PostTime) & " " & Hour(PostTime) & ":" & Minute(PostTime) & ":" & Second(PostTime)

			If ZC_POST_STATIC_MODE<>"STATIC" Then FullUrl=Replace(Url,BlogHost,"<#ZC_BLOG_HOST#>")

		End If

		FUrl=""

		LoadInfoByArray=True

		Call Filter_Plugin_TArticle_LoadInfoByArray(ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,FullUrl,FType,MetaString)

	End Function



	Public Function Post()

		Call Filter_Plugin_TArticle_Post(ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,FullUrl,FType,MetaString)

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

		Dim sTitle,sIntro,sContent
		sTitle=Title
		sIntro=Content
		sContent=Content
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

		Alias=TransferHTML(Alias,"[directory&file]")
		If Left(Alias,1)="/" Then Alias=Right(Alias,Len(Alias)-1)
		If Right(Alias,1)="/" Then Alias=Left(Alias,Len(Alias)-1)
		Alias=FilterSQL(Alias)

		Call GetUsersbyUserIDList(AuthorID)

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
		'If Len(Intro)=0 Then Intro=Left(Content,ZC_ARTICLE_EXCERPT_MAX) & "..."

		TemplateName=UCase(FilterSQL(TemplateName))

		If IsPage=False Then
			If Categorys(CateID).LogTemplate<>"" Then
				If TemplateName=Categorys(CateID).LogTemplate Then TemplateName=""
			Else
				If TemplateName="SINGLE" Then TemplateName=""
			End If
		Else
			If TemplateName="PAGE" Then TemplateName=""
		End If

		Dim objRS
		If ID=0 Then
			objConn.Execute("INSERT INTO [blog_Article]([log_CateID],[log_AuthorID],[log_Level],[log_Title],[log_Intro],[log_Content],[log_PostTime],[log_IP],[log_Tag],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_ViewNums],[log_Type],[log_Meta]) VALUES ("&CateID&","&AuthorID&","&Level&",'"&Title&"','"&Intro&"','"&Content&"','"&PostTime&"','"&IP&"','"&Tag&"','"&Alias&"',"&CLng(Istop)&",'"&TemplateName&"','"&FullUrl&"',0,"&CLng(FType)&",'"&MetaString&"')")
			Set objRS=objConn.Execute("SELECT MAX([log_ID]) FROM [blog_Article]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
			End If
			Set objRS=Nothing

			'FullUrl=Replace(Url,BlogHost,"<#ZC_BLOG_HOST#>")
			'objConn.Execute("UPDATE [blog_Article] SET [log_FullUrl]='"&FullUrl&"' WHERE [log_ID] =" & ID)
		Else
			Set objRS=objConn.Execute("SELECT [log_CateID],[log_AuthorID] FROM [blog_Article] WHERE [log_ID] =" & ID)
			If (Not objRS.bof) And (Not objRS.eof) Then
				If objRS(0)<>CateID Then
					Categorys(CateID).ExID=objRS(0)
				End If
				If objRS(1)<>AuthorID Then
					Users(AuthorID).ExID=objRS(1)
				End If
			End If
			Set objRS=Nothing

			'FullUrl=Replace(Url,BlogHost,"<#ZC_BLOG_HOST#>")

			objConn.Execute("UPDATE [blog_Article] SET [log_CateID]="&CateID&",[log_AuthorID]="&AuthorID&",[log_Level]="&Level&",[log_Title]='"&Title&"',[log_Intro]='"&Intro&"',[log_Content]='"&Content&"',[log_PostTime]='"&PostTime&"',[log_IP]='"&IP&"',[log_Tag]='"&Tag&"',[log_Url]='"&Alias&"',[log_Istop]="&CLng(Istop)&",[log_Template]='"&TemplateName&"',[log_FullUrl]='"&FullUrl&"',[log_Type]="&CLng(FType)&",[log_Meta]='"&MetaString&"' WHERE [log_ID] =" & ID)
		End If

		Title=sTitle
		Content=sIntro
		Content=sContent

		Post=True

	End Function



	Public Function Export_Tag

		If Disable_Export_Tag=True Then Exit Function

		'plugin node
		bAction_Plugin_TArticle_Export_Tag_Begin=False
		For Each sAction_Plugin_TArticle_Export_Tag_Begin in Action_Plugin_TArticle_Export_Tag_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_Tag_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_Tag_Begin)
			If bAction_Plugin_TArticle_Export_Tag_Begin=True Then Exit Function
		Next

		Call GetTagsbyTagIDList(Tag)

		'Tag
		Dim t,i,s,j

		If Tag<>"" Then

			HasTag=True

			s=Replace(Tag,"}","")
			t=Split(s,"{")
			For i=LBound(t) To UBound(t)
				If t(i)<>"" Then

					If IsObject(Tags(t(i)))=True Then
						j=GetTemplate("TEMPLATE_B_ARTICLE_TAG")

						Template_Article_Tag=Template_Article_Tag & Tags(t(i)).MakeTemplate(j)
					End If
				End If
			Next

		End If

		Template_Article_Tag=TransferHTML(Template_Article_Tag,"[anti-zc_blog_host]")

		Export_Tag=True

	End Function




	Function Export_CMTandTB(intPage)

		If Disable_Export_CMTandTB=True Then Exit Function

		If ZC_COMMENT_TURNOFF=True Then
			Template_Article_Comment=""
			Exit Function
		End If

		Call CheckParameter(intPage,"int",1)

		'plugin node
		bAction_Plugin_TArticle_Export_CMTandTB_Begin=False
		For Each sAction_Plugin_TArticle_Export_CMTandTB_Begin in Action_Plugin_TArticle_Export_CMTandTB_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_CMTandTB_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_CMTandTB_Begin)
			If bAction_Plugin_TArticle_Export_CMTandTB_Begin=True Then Exit Function
		Next

		Dim intCommnums
		intCommnums=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Comment] WHERE [log_ID] =" & ID & " AND [comm_isCheck]=0 AND [comm_ParentID]=0")(0)
		If intCommnums > 0 Then

			HasCMTandTB=True

			Dim strC_Count,strC,strT_Count,strT

			Dim objComment
			Dim objTrackBack

			Dim i,j,s,t

			Dim comments_ID()
			Dim comments_ParentID()
			Dim comments_Template()


			Dim IDandTemp
			Set IDandTemp = CreateObject("Scripting.Dictionary")

			Dim tree
			Set tree = CreateObject("Scripting.Dictionary")

			Dim all
			Set all = CreateObject("Scripting.Dictionary")


			Dim alltemplate
			Set alltemplate = CreateObject("Scripting.Dictionary")

			Dim objRS

			Dim order
			order=IIf(ZC_COMMENT_REVERSE_ORDER_EXPORT,"DESC","ASC")

			strC_Count=0
			Set objRS=Server.CreateObject("ADODB.Recordset")
			objRS.CursorType = adOpenKeyset
			objRS.LockType = adLockReadOnly
			objRS.ActiveConnection=objConn
			ZC_COMMENTS_DISPLAY_COUNT=IIF(ZC_COMMENTS_DISPLAY_COUNT=0,10000,ZC_COMMENTS_DISPLAY_COUNT)
			Dim PageSize2,PageSize3
			PageSize2=ZC_COMMENTS_DISPLAY_COUNT
			PageSize3=ZC_COMMENTS_DISPLAY_COUNT
			If intPage<1 Then intPage=1
			If PageSize3*intPage>intCommnums Then
				PageSize2=CLng(intCommnums Mod PageSize3)
				If PageSize3*(intPage-1)+PageSize2>intCommnums Then
					Template_Article_Comment="<ins style=""display:none;"" id=""AjaxCommentBegin""></ins>" & Template_Article_Comment & "<ins style=""display:none;"" id=""AjaxCommentEnd""></ins>"
					Export_CMTandTB=True
					Exit Function
				End If
			End If
			j=PageSize2

			objRS.Source="SELECT * FROM (SELECT TOP "&PageSize2&" *  FROM (SELECT TOP "&(PageSize3*intPage)&" * FROM [blog_Comment]  WHERE ([log_ID]="&id&" AND [comm_isCheck]=0 AND [comm_ParentID]=0) ORDER BY [comm_id] "&order&") As [Test] ORDER BY [comm_id] "&IIf(ZC_COMMENT_REVERSE_ORDER_EXPORT,"ASC","DESC")&" ) As [test] order by [comm_posttime] "&order
			objRS.Open()
			j=PageSize2
			Dim intPageAll
			If (intCommnums Mod PageSize3)=0 Then
				intPageAll=Int(intCommnums/PageSize3)
			Else
				intPageAll=Int(intCommnums/PageSize3)+1
			End If


			If (Not objRS.bof) And (Not objRS.eof) Then


				strC=GetTemplate("TEMPLATE_B_ARTICLE_COMMENT")
				For i=1 To j


					Set objComment=New TComment
					objComment.LoadInfoByArray(Array(objRS("comm_ID"),objRS("log_ID"),objRS("comm_AuthorID"),objRS("comm_Author"),objRS("comm_Content"),objRS("comm_Email"),objRS("comm_HomePage"),objRS("comm_PostTime"),objRS("comm_IP"),objRS("comm_Agent"),objRS("comm_Reply"),objRS("comm_LastReplyIP"),objRS("comm_LastReplyTime"),objRS("comm_ParentID"),objRS("comm_IsCheck"),objRs("comm_Meta")))

					Call GetUsersbyUserIDList(objRS("comm_AuthorID"))

					'objComment.Count=0
					objComment.Count=IIf(ZC_COMMENT_REVERSE_ORDER_EXPORT,intCommnums-((intPage-1)*ZC_COMMENTS_DISPLAY_COUNT+i)+1,(intPage-1)*ZC_COMMENTS_DISPLAY_COUNT+i)

					tree.add objComment.ID, objComment.MakeTemplate(strC)'objComment

					If ZC_COMMENT_REVERSE_ORDER_EXPORT Then
						t=objRS("comm_ID")
					ElseIf i=1 Then
						t=objRS("comm_ID")
					End If

					Set objComment=Nothing

					objRS.MoveNext
					If objRS.eof Then Exit For

				Next

				Dim objRS2

				Set objRS2=objConn.Execute("SELECT * FROM [blog_Comment] WHERE ([log_ID]=" & ID &" AND [comm_isCheck]=0 AND [comm_ParentID]<>0 And [comm_ID]>"&t&")  ORDER BY [comm_PostTime] DESC")
				If (Not objRS2.bof) And (Not objRS2.eof) Then
					Do While Not objRS2.eof

						Set objComment=New TComment
						objComment.LoadInfoByArray(Array(objRS2("comm_ID"),objRS2("log_ID"),objRS2("comm_AuthorID"),objRS2("comm_Author"),objRS2("comm_Content"),objRS2("comm_Email"),objRS2("comm_HomePage"),objRS2("comm_PostTime"),objRS2("comm_IP"),objRS2("comm_Agent"),objRS2("comm_Reply"),objRS2("comm_LastReplyIP"),objRS2("comm_LastReplyTime"),objRS2("comm_ParentID"),objRS2("comm_IsCheck"),objRS2("comm_Meta")))
						'Call GetUsersbyUserIDList(objRS2("comm_AuthorID"))
						objComment.Count=0
						all.add objComment.ID, objComment.ParentID
						alltemplate.add objComment.ID,objComment.MakeTemplate(strC)
						Set objComment=Nothing
						objRS2.MoveNext
					Loop
				End If
				objRS2.Close
				Set objRS2=Nothing

			End if

			objRS.Close()
			Set objRS=Nothing


			For Each s In tree.Keys
				t="<!--rev"&s&"-->"
				If SearchChildCommentsInDic(s,t,all,alltemplate)=True Then
					tree.Item(s) =Replace(tree.Item(s),"<!--rev"&s&"-->",t)
				End If
				t=""

			Next

			'输出树
			For Each s In tree.Items
				'If ZC_COMMENT_REVERSE_ORDER_EXPORT=True Then
					Template_Article_Comment=Template_Article_Comment & s
				'Else
				'	Template_Article_Comment=s & Template_Article_Comment
				'End If
			Next

		End If

		Call ExportCMTandTBBar(intPage,intPageAll)

		Template_Article_Comment="<ins style=""display:none;"" id=""AjaxCommentBegin""></ins>" & Template_Article_Comment & Template_Article_Comment_Pagebar &"<ins style=""display:none;"" id=""AjaxCommentEnd""></ins>"

		Template_Article_Comment=Replace(Template_Article_Comment,"<!--(count-->0<!--count)-->","<span class=""revcount""></span>")
		Template_Article_Comment=Replace(Template_Article_Comment,"<!--(count-->","")
		Template_Article_Comment=Replace(Template_Article_Comment,"<!--count)-->","")

		Template_Article_Comment=TransferHTML(Template_Article_Comment,"[anti-zc_blog_host]")

		Export_CMTandTB=True

	End Function




	Public Function ExportCMTandTBBar(intPage,intPageAll)

		'plugin node
		bAction_Plugin_TArticle_ExportCMTandTBBar_Begin=False
		For Each sAction_Plugin_TArticle_ExportCMTandTBBar_Begin in Action_Plugin_TArticle_ExportCMTandTBBar_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_ExportCMTandTBBar_Begin) Then Call Execute(sAction_Plugin_TArticle_ExportCMTandTBBar_Begin)
			If bAction_Plugin_TArticle_ExportCMTandTBBar_Begin=True Then Exit Function
		Next

		Dim s,l,r

		If intPageAll>1 Then
			Dim lp,rp

			lp="GetComments("&ID&","&(intPage-1)&")"
			rp="GetComments("&ID&","&(intPage+1)&")"

			l=GetTemplate("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR_L")
			r=GetTemplate("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR_R")

			l=Replace(l,"<#article/comment_pagebar_l/url#>",lp)
			r=Replace(r,"<#article/comment_pagebar_r/url#>",rp)

			If intPage=1 Then
				l=""
			End If
			If intPage=intPageAll Then
				r=""
			End If

			s=GetTemplate("TEMPLATE_B_ARTICLE_COMMENT_PAGEBAR")
			s=Replace(s,"<#template:pagebar#>",l&r&"")
		End If

		Template_Article_Comment_Pagebar=s

		ExportCMTandTBBar=True

		'plugin node
		bAction_Plugin_TArticle_ExportCMTandTBBar_End=False
		For Each sAction_Plugin_TArticle_ExportCMTandTBBar_End in Action_Plugin_TArticle_ExportCMTandTBBar_End
			If Not IsEmpty(sAction_Plugin_TArticle_ExportCMTandTBBar_End) Then Call Execute(sAction_Plugin_TArticle_ExportCMTandTBBar_End)
			If bAction_Plugin_TArticle_ExportCMTandTBBar_End=True Then Exit Function
		Next

	End Function




	Private Function SearchChildCommentsInDic(ByVal id,ByRef t,ByVal alltree,ByVal alltemplate)

		Dim s

		For Each s In alltree.Keys

			If alltree.Item(s)=id Then
				t=Replace(t,"<!--rev"&id&"-->","<!--rev"&id&"-->" & alltemplate.Item(CLng(s) ) )
				Call SearchChildCommentsInDic(s,t,alltree,alltemplate)
				SearchChildCommentsInDic=True
			End If

		Next

	End Function




	Function Export_NavBar()

		If Disable_Export_NavBar=True Then Exit Function

		'plugin node
		bAction_Plugin_TArticle_Export_NavBar_Begin=False
		For Each sAction_Plugin_TArticle_Export_NavBar_Begin in Action_Plugin_TArticle_Export_NavBar_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_NavBar_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_NavBar_Begin)
			If bAction_Plugin_TArticle_Export_NavBar_Begin=True Then Exit Function
		Next


		If ZC_USE_NAVIGATE_ARTICLE=False Or IsPage=True Then

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

		Set objRS=objConn.Execute("SELECT TOP 1 [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_Level]>2) AND ([log_PostTime]<" & ZC_SQL_POUND_KEY & PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] DESC")
		If (Not objRS.bof) And (Not objRS.eof) Then

			Set objNavArticle=New TArticle
			If objNavArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then
				strName=objNavArticle.Title
				strUrl=objNavArticle.Url
			End If
			Set objNavArticle=Nothing

			s=GetTemplate("TEMPLATE_B_ARTICLE_NAVBAR_L")

			's=Replace(s,"<#article/nav_l/url#>","<#ZC_BLOG_HOST#>view.asp?navp="&ID)
			's=Replace(s,"<#article/nav_l/name#>",ZC_MSG146)
			s=Replace(s,"<#article/nav_l/url#>",strUrl)
			s=Replace(s,"<#article/nav_l/name#>",strName)

			Template_Article_Navbar_L=s

		End If
		Set objRS=Nothing

		Set objRS=objConn.Execute("SELECT TOP 1 [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_Level]>2) AND ([log_PostTime]>" & ZC_SQL_POUND_KEY & PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] ASC")

		If (Not objRS.bof) And (Not objRS.eof) Then

			Set objNavArticle=New TArticle
			If objNavArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then
				strName=objNavArticle.Title
				strUrl=objNavArticle.Url
			End If
			Set objNavArticle=Nothing

			t=GetTemplate("TEMPLATE_B_ARTICLE_NAVBAR_R")

			't=Replace(t,"<#article/nav_r/url#>","<#ZC_BLOG_HOST#>view.asp?navn="&ID)
			't=Replace(t,"<#article/nav_r/name#>",ZC_MSG148)
			t=Replace(t,"<#article/nav_r/url#>",strUrl)
			t=Replace(t,"<#article/nav_r/name#>",strName)

			Template_Article_Navbar_R=t

		End If
		Set objRS=Nothing

		Export_NavBar=True

	End Function


	Function Export_CommentPost()

		If Disable_Export_CommentPost=True Then Exit Function

		'plugin node
		bAction_Plugin_TArticle_Export_CommentPost_Begin=False
		For Each sAction_Plugin_TArticle_Export_CommentPost_Begin in Action_Plugin_TArticle_Export_CommentPost_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_CommentPost_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_CommentPost_Begin)
			If bAction_Plugin_TArticle_Export_CommentPost_Begin=True Then Exit Function
		Next

		If ZC_COMMENT_TURNOFF=True Then
			Template_Article_Commentpost=""
			Exit Function
		End If

		If Level<4 Then Exit Function

		Template_Article_Commentpost=GetTemplate("TEMPLATE_B_ARTICLE_COMMENTPOST")

		Dim RE
		Set RE = New RegExp
		RE.IgnoreCase = True
		RE.Global = True

		If InStr(Template_Article_Commentpost,"template:article_commentpost-verify:")>0  Then'2.2新增条件判断
			If ZC_COMMENT_VERIFY_ENABLE=True Then
				Template_Article_Commentpost=Replace(Template_Article_Commentpost,"template:article_commentpost-verify:begin","")
				Template_Article_Commentpost=Replace(Template_Article_Commentpost,"template:article_commentpost-verify:end","")
			Else
				RE.Pattern = "<#template:article_commentpost-verify:begin#>(.|\n)*<#template:article_commentpost-verify:end#>"
				Template_Article_Commentpost = RE.Replace(Template_Article_Commentpost, "")
			End If
		Else
			If ZC_COMMENT_VERIFY_ENABLE=True Then
				Template_Article_Commentpost_Verify=GetTemplate("TEMPLATE_B_ARTICLE_COMMENTPOST-VERIFY")
			End If
			Template_Article_Commentpost=Replace(Template_Article_Commentpost,"<#template:article_commentpost-verify#>",Template_Article_Commentpost_Verify)
		End If


	End Function


	'相关文章的生成
	Function Export_Mutuality()

		If Disable_Export_Mutuality=True Then Exit Function

		'plugin node
		bAction_Plugin_TArticle_Export_Mutuality_Begin=False
		For Each sAction_Plugin_TArticle_Export_Mutuality_Begin in Action_Plugin_TArticle_Export_Mutuality_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_Mutuality_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_Mutuality_Begin)
			If bAction_Plugin_TArticle_Export_Mutuality_Begin=True Then Exit Function
		Next

		If ZC_MUTUALITY_COUNT=0 Then
			Export_Mutuality=True
			Exit Function
		End If

		If Tag<>"" Then

			Dim strCC_Count,strCC_ID,strCC_Name,strCC_Url,strCC_PostTime,strCC_Title
			Dim strCC
			Dim i,j,s
			Dim objRS
			Dim strSQL

			Set objRS=Server.CreateObject("ADODB.Recordset")

			strSQL="SELECT TOP "& ZC_MUTUALITY_COUNT &" [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_Level]>2) AND [log_ID]<>"& ID
			strSQL = strSQL & " AND ("

			Dim aryTAGs
			s=Replace(Tag,"}","")
			aryTAGs=Split(s,"{")

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

					HasMutuality=True

					Set objArticle=New TArticle

					If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17)))  Then

						strCC_Count=strCC_Count+1
						strCC_ID=objArticle.ID
						strCC_Url=objArticle.Url
						strCC_PostTime=objArticle.PostTime
						strCC_Title=objArticle.Title

						strCC=GetTemplate("TEMPLATE_B_ARTICLE_MUTUALITY")

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


	Public Function Export(intType)

		'plugin node
		bAction_Plugin_TArticle_Export_Begin=False
		For Each sAction_Plugin_TArticle_Export_Begin in Action_Plugin_TArticle_Export_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Export_Begin) Then Call Execute(sAction_Plugin_TArticle_Export_Begin)
			If bAction_Plugin_TArticle_Export_Begin=True Then Exit Function
		Next

		If IsEmpty(html)=True Then html=Template

		Call GetUsersbyUserIDList(AuthorID)

		If (ZC_DISPLAY_MODE_INTRO=intType) Or (ZC_DISPLAY_MODE_ONTOP=intType) Or (ZC_DISPLAY_MODE_SEARCH=intType) Or (ZC_DISPLAY_MODE_SYSTEMPAGE=intType) Or (ZC_DISPLAY_MODE_COMMENTS=intType) Then
			Disable_Export_Tag=False
			Disable_Export_CMTandTB=True
			Disable_Export_CommentPost=True
			Disable_Export_Mutuality=True
			Disable_Export_NavBar=True
			If ZC_DISPLAY_MODE_SYSTEMPAGE=intType Then
				Disable_Export_Tag=True
			End If
			If ZC_DISPLAY_MODE_ONTOP=intType Then
				'Disable_Export_Tag=True
				subhtml_TemplateName=""
				subhtml=GetTemplate("TEMPLATE_B_ARTICLE-ISTOP")
			End If
			If ZC_DISPLAY_MODE_SEARCH=intType Then
				Disable_Export_Tag=True
				subhtml_TemplateName=""
				subhtml=GetTemplate("TEMPLATE_B_ARTICLE-SEARCH-CONTENT")
			End If
			If ZC_DISPLAY_MODE_COMMENTS=intType Then
				Disable_Export_Tag=True
				Disable_Export_CMTandTB=False
				subhtml_TemplateName=""
				subhtml="<#template:article_comment#>"
			End If
		End If

		If ZC_DISPLAY_MODE_ALL=intType And IsPage=True Then
			Disable_Export_CMTandTB=False
			Disable_Export_CommentPost=False
			Disable_Export_Tag=True
			Disable_Export_Mutuality=True
			Disable_Export_NavBar=True
		End If

		Call Export_Tag
		Call Export_CMTandTB(CommentsPage)
		Call Export_CommentPost
		Call Export_Mutuality
		Call Export_NavBar

		Dim RE ,Match,Matches
		Set RE = New RegExp
			RE.Pattern = "\<\#template\:(article\-([a-z0-9]*)([\-a-z0-9]*))\#\>"
			RE.IgnoreCase = True
			RE.Global = True
			Set Matches = RE.Execute(html)
			For Each Match in Matches
				If IsEmpty(subhtml_TemplateName) Then subhtml_TemplateName="template:"&Match.SubMatches(0)&""
				If IsEmpty(subhtml) Then subhtml=GetTemplate("TEMPLATE_B_"& UCase(Match.SubMatches(0)))
				Exit For
			Next
			Set Matches = Nothing
		Set RE = Nothing

		If subhtml="" Then
			subhtml=html
			subhtml_TemplateName="template:article_single"
			html="<#template:article_single#>"
		End If


		'plugin node
		Call Filter_Plugin_TArticle_Export_Template(html,subhtml)

		'plugin node
		Call Filter_Plugin_TArticle_Export_Template_Sub(Template_Article_Comment,Template_Article_Trackback,Template_Article_Tag,Template_Article_Commentpost,Template_Article_Navbar_L,Template_Article_Navbar_R,Template_Article_Mutuality)

		subhtml=Replace(subhtml,"<#template:article_comment#>",Template_Article_Comment)
		subhtml=Replace(subhtml,"<#template:article_trackback#>",Template_Article_Trackback)
		subhtml=Replace(subhtml,"<#template:article_comment_pagebar#>","")
		subhtml=Replace(subhtml,"<#template:article_commentpost#>",Template_Article_Commentpost)
		subhtml=Replace(subhtml,"<#template:article_tag#>",Template_Article_Tag)
		subhtml=Replace(subhtml,"<#template:article_navbar_l#>",Template_Article_Navbar_L)
		subhtml=Replace(subhtml,"<#template:article_navbar_r#>",Template_Article_Navbar_R)
		subhtml=Replace(subhtml,"<#template:article_mutuality#>",Template_Article_Mutuality)

		Dim da
		Dim s,t
		Set da=CreateObject("Scripting.Dictionary")

		Set RE = New RegExp
			RE.Pattern = "<#article/((category/meta|author/meta|meta)/([a-z0-9_]{1,}))#>"
			RE.IgnoreCase = True
			RE.Global = True
			Set Matches = RE.Execute(html&subhtml)
			For Each Match in Matches
				s=Match.SubMatches(0)
				If da.Exists(s)=False Then da.add s,s
			Next
			Set Matches = Nothing
		Set RE = Nothing

		s=da.Items
		For t = 0 To da.Count -1
			If InStr(s(t),"category/meta/")>0 Then
				s(t)=Replace(s(t),"category/meta/","")
				subhtml=Replace(subhtml,"<#article/category/meta/" & s(t) & "#>",Categorys(CateID).Meta.GetValue(s(t)) )
				html = Replace(html,"<#article/category/meta/" & s(t) & "#>",Categorys(CateID).Meta.GetValue(s(t)) )
			ElseIf InStr(s(t),"author/meta/")>0 Then
				s(t)=Replace(s(t),"author/meta/","")
				subhtml=Replace(subhtml,"<#article/author/meta/" & s(t) & "#>",Users(AuthorID).Meta.GetValue(s(t)) )
				html = Replace(html,"<#article/author/meta/" & s(t) & "#>",Users(AuthorID).Meta.GetValue(s(t)) )
			Else
				s(t)=Replace(s(t),"meta/","")
				subhtml=Replace(subhtml,"<#article/meta/" & s(t) & "#>",Meta.GetValue(s(t)) )
				html = Replace(html,"<#article/meta/" & s(t) & "#>",Meta.GetValue(s(t)) )
			End If
		Next

		If ZC_MULTI_DOMAIN_SUPPORT=True And ZC_PERMANENT_DOMAIN_ENABLE=False Then
			Content=Replace(Content,"href=""" & BlogHost,"href=""<#ZC_BLOG_HOST#>")
			Content=Replace(Content,"src=""" & BlogHost,"src=""<#ZC_BLOG_HOST#>")
			Intro=Replace(Intro,"href=""" & BlogHost,"href=""<#ZC_BLOG_HOST#>")
			Intro=Replace(Intro,"src=""" & BlogHost,"src=""<#ZC_BLOG_HOST#>")
		End If

		Dim aryTemplateTagsName()
		Dim aryTemplateTagsValue()
		Dim i,j
		ReDim aryTemplateTagsName(66)
		ReDim aryTemplateTagsValue(66)

		aryTemplateTagsName(1)="article/id"
		aryTemplateTagsValue(1)=ID
		aryTemplateTagsName(2)="article/level"
		aryTemplateTagsValue(2)=Level
		aryTemplateTagsName(3)="article/title"
		If intType=ZC_DISPLAY_MODE_SEARCH Then
			aryTemplateTagsValue(3)=Search(Title,SearchText)
		Else
			aryTemplateTagsValue(3)=HtmlTitle
		End If
		aryTemplateTagsName(4)="article/intro"
		If intType=ZC_DISPLAY_MODE_SEARCH Then
			'aryTemplateTagsValue(4)=Search(TransferHTML(Intro & Content,"[html-format]"),Request.QueryString("q"))
			aryTemplateTagsValue(4)=Trim(Search(TransferHTML(Intro & Content,"[nohtml]"),SearchText))
		Else
			If Level=2 Then
				aryTemplateTagsValue(4)="<p>"&ZC_MSG043&"</p>"
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
		aryTemplateTagsValue(11)=TransferHTML(HtmlUrl,"[anti-zc_blog_host]")

		aryTemplateTagsName(12)="article/category/id"
		aryTemplateTagsValue(12)=Categorys(CateID).ID
		aryTemplateTagsName(13)="article/category/name"
		aryTemplateTagsValue(13)=Categorys(CateID).HtmlName
		aryTemplateTagsName(15)="article/category/order"
		aryTemplateTagsValue(15)=Categorys(CateID).Order
		aryTemplateTagsName(16)="article/category/count"
		aryTemplateTagsValue(16)=Categorys(CateID).Count
		aryTemplateTagsName(17)="article/category/url"
		aryTemplateTagsValue(17)=TransferHTML(Categorys(CateID).HtmlUrl,"[anti-zc_blog_host]")

		aryTemplateTagsName(18)="article/author/id"
		aryTemplateTagsValue(18)=Users(AuthorID).ID
		aryTemplateTagsName(19)="article/author/name"
		aryTemplateTagsValue(19)=Users(AuthorID).FirstName
		aryTemplateTagsName(20)="article/author/level"
		aryTemplateTagsValue(20)=Users(AuthorID).Level
		aryTemplateTagsName(21)="article/author/email"
		aryTemplateTagsValue(21)=Users(AuthorID).Email
		aryTemplateTagsName(22)="article/author/homepage"
		aryTemplateTagsValue(22)=Users(AuthorID).HomePage
		aryTemplateTagsName(23)="article/author/count"
		aryTemplateTagsValue(23)=Users(AuthorID).Count
		aryTemplateTagsName(24)="article/author/url"
		aryTemplateTagsValue(24)=TransferHTML(Users(AuthorID).HtmlUrl,"[anti-zc_blog_host]")

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
		aryTemplateTagsValue(30)=Right("0"&Month(PostTime),2)
		aryTemplateTagsName(31)="article/posttime/monthname"
		aryTemplateTagsValue(31)=ZVA_Month(Month(PostTime))
		aryTemplateTagsName(32)="article/posttime/day"
		aryTemplateTagsValue(32)=Right("0"&Day(PostTime),2)
		aryTemplateTagsName(33)="article/posttime/weekday"
		aryTemplateTagsValue(33)=Weekday(PostTime)
		aryTemplateTagsName(34)="article/posttime/weekdayname"
		aryTemplateTagsValue(34)=ZVA_Week(Weekday(PostTime))
		aryTemplateTagsName(35)="article/posttime/hour"
		aryTemplateTagsValue(35)=Right("0"&Hour(PostTime),2)
		aryTemplateTagsName(36)="article/posttime/minute"
		aryTemplateTagsValue(36)=Right("0"&Minute(PostTime),2)
		aryTemplateTagsName(37)="article/posttime/second"
		aryTemplateTagsValue(37)=Right("0"&Second(PostTime),2)

		aryTemplateTagsName(38)="article/commentrss"
		aryTemplateTagsValue(38)=TransferHTML(WfwCommentRss,"[anti-zc_blog_host]")
		aryTemplateTagsName(39)="article/commentposturl"
		aryTemplateTagsValue(39)=TransferHTML(CommentPostUrl,"[html-format][anti-zc_blog_host]")
		aryTemplateTagsName(40)="article/pretrackback_url"
		aryTemplateTagsValue(40)=TransferHTML(PreTrackBack,"[html-format][anti-zc_blog_host]")
		aryTemplateTagsName(41)="article/trackbackkey"
		aryTemplateTagsValue(41)=TrackBackKey
		aryTemplateTagsName(42)="article/commentkey"
		aryTemplateTagsValue(42)=CommentKey

		aryTemplateTagsName(43)="article/staticname"
		aryTemplateTagsValue(43)=StaticName
		aryTemplateTagsName(44)="article/category/staticname"
		aryTemplateTagsValue(44)=Categorys(CateID).StaticName
		aryTemplateTagsName(45)="article/author/staticname"
		aryTemplateTagsValue(45)=Users(AuthorID).StaticName
		aryTemplateTagsName(46)="article/tagtoname"
		aryTemplateTagsValue(46)=TagToName
		aryTemplateTagsName(47)="article/firsttagintro"
		aryTemplateTagsValue(47)=FirstTagIntro

		aryTemplateTagsName(48)="article/posttime/monthnameabbr"
		aryTemplateTagsValue(48)=ZVA_Month_Abbr(Month(PostTime))
		aryTemplateTagsName(49)="article/posttime/weekdaynameabbr"
		aryTemplateTagsValue(49)=ZVA_Week_Abbr(Weekday(PostTime))
		aryTemplateTagsName(50)="article/author/intro"
		aryTemplateTagsValue(50)=Users(AuthorID).Intro

		aryTemplateTagsName(51)="template:sidebar"
		aryTemplateTagsValue(51)=GetTemplate("CACHE_SIDEBAR")
		aryTemplateTagsName(52)="template:sidebar2"
		aryTemplateTagsValue(52)=GetTemplate("CACHE_SIDEBAR2")
		aryTemplateTagsName(53)="template:sidebar3"
		aryTemplateTagsValue(53)=GetTemplate("CACHE_SIDEBAR3")
		aryTemplateTagsName(54)="template:sidebar4"
		aryTemplateTagsValue(54)=GetTemplate("CACHE_SIDEBAR4")
		aryTemplateTagsName(55)="template:sidebar5"
		aryTemplateTagsValue(55)=GetTemplate("CACHE_SIDEBAR5")

		aryTemplateTagsName(56)="article/alias"
		aryTemplateTagsValue(56)=Alias
		aryTemplateTagsName(57)="article/loadviewcount"
		' aryTemplateTagsValue(57)="<span id=""spn"&ID&"""></span><script type=""text/javascript"">LoadViewCount("&ID&")</script>"
		aryTemplateTagsValue(57)="<span class=""LoadView"" id=""spn"&ID&""" data-id="""&ID&"""></span>"
		aryTemplateTagsName(58)="article/addviewcount"
		' aryTemplateTagsValue(58)="<span id=""spn"&ID&"""></span><script type=""text/javascript"">AddViewCount("&ID&")</script>"
		aryTemplateTagsValue(58)="<span class=""AddView"" id=""spn"&ID&""" data-id="""&ID&"""></span>"


		aryTemplateTagsName(59)="article/category/parent/id"
		aryTemplateTagsName(60)="article/category/parent/name"
		aryTemplateTagsName(61)="article/category/parent/order"
		aryTemplateTagsName(62)="article/category/parent/count"
		aryTemplateTagsName(63)="article/category/parent/url"
		aryTemplateTagsName(64)="article/category/parent/staticname"

		If Categorys(CateID).ParentID<>0 Then
		aryTemplateTagsValue(59)=Categorys(Categorys(CateID).ParentID).ID
		aryTemplateTagsValue(60)=Categorys(Categorys(CateID).ParentID).HtmlName
		aryTemplateTagsValue(61)=Categorys(Categorys(CateID).ParentID).Order
		aryTemplateTagsValue(62)=Categorys(Categorys(CateID).ParentID).Count
		aryTemplateTagsValue(63)=TransferHTML(Categorys(Categorys(CateID).ParentID).HtmlUrl,"[anti-zc_blog_host]")
		aryTemplateTagsValue(64)=Categorys(Categorys(CateID).ParentID).StaticName
		End If

		aryTemplateTagsName(65)="article/author/avatar"
		aryTemplateTagsValue(65)=Users(AuthorID).Avatar
		aryTemplateTagsName(66)="article/author/levelname"
		aryTemplateTagsValue(66)=Users(AuthorID).LevelName


		Call Filter_Plugin_TArticle_Export_TemplateTags(aryTemplateTagsName,aryTemplateTagsValue)

		j=UBound(aryTemplateTagsName)
		For i=1 to j
			If IsNull(aryTemplateTagsValue(i))=False Then
				subhtml=Replace(subhtml,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
				html = Replace(html,"<#" & aryTemplateTagsName(i) & "#>", aryTemplateTagsValue(i))
			End If
		Next

		Set RE = New RegExp
		RE.IgnoreCase = True
		RE.Global = True

		If HasTag=False Then
			RE.Pattern = "<#template:article_tag:begin#>(.|\n)*<#template:article_tag:end#>"
			subhtml = RE.Replace(subhtml, "")
		Else
			subhtml=Replace(subhtml,"<#template:article_tag:begin#>","")
			subhtml=Replace(subhtml,"<#template:article_tag:end#>","")
		End If
		If HasCMTandTB=False Then
			RE.Pattern = "<#template:article_comment:begin#>(.|\n)*<#template:article_comment:end#>"
			subhtml = RE.Replace(subhtml, "<ins style=""display:none;"" id=""AjaxCommentEnd""></ins><ins style=""display:none;"" id=""AjaxCommentBegin""></ins>")
		Else
			subhtml=Replace(subhtml,"<#template:article_comment:begin#>","")
			subhtml=Replace(subhtml,"<#template:article_comment:end#>","")
		End IF
		If HasMutuality=False Then
			RE.Pattern = "<#template:article_mutuality:begin#>(.|\n)*<#template:article_mutuality:end#>"
			subhtml = RE.Replace(subhtml, "")
		Else
			subhtml=Replace(subhtml,"<#template:article_mutuality:begin#>","")
			subhtml=Replace(subhtml,"<#template:article_mutuality:end#>","")
		End If

		If Categorys(CateID).ParentID=0 Then
			RE.Pattern = "<#template:article_category_parent:begin#>(.|\n)*<#template:article_category_parent:end#>"
			subhtml = RE.Replace(subhtml, "")
		Else
			subhtml=Replace(subhtml,"<#template:article_category_parent:begin#>","")
			subhtml=Replace(subhtml,"<#template:article_category_parent:end#>","")
		End If
		Set RE = Nothing


		html=Replace(html,"<#"&subhtml_TemplateName&"#>",subhtml)

		Export=True

		'plugin node
		bAction_Plugin_TArticle_Export_End=False
		For Each sAction_Plugin_TArticle_Export_End in Action_Plugin_TArticle_Export_End
			If Not IsEmpty(sAction_Plugin_TArticle_Export_End) Then Call Execute(sAction_Plugin_TArticle_Export_End)
			If bAction_Plugin_TArticle_Export_End=True Then Exit Function
		Next

	End Function




	Function Build()

		Dim aryTemplateTagsName
		Dim aryTemplateTagsValue

		Dim i,j

		Call Filter_Plugin_TArticle_Build_Template(html)

		TemplateTagsDic.Item("BlogTitle")=HtmlTitle
		TemplateTagsDic.Item("ZC_BLOG_HOST")=BlogHost

		If Template_Calendar="" Then Template_Calendar="<script src=""<#ZC_BLOG_HOST#>zb_system/function/c_html_js.asp?date=now"" type=""text/javascript""></script>"

		If ZC_MULTI_DOMAIN_SUPPORT=True And ZC_PERMANENT_DOMAIN_ENABLE=False Then
			Dim x,y
			x=CStr(Replace(Url,BlogHost,""))
			x=Replace(x,"/////","/"):x=Replace(x,"////","/"):x=Replace(x,"///","/"):x=Replace(x,"//","/")
			For i=1 To UBound(Split(x,"/"))
				y=y & "../"
			Next
			If y="" Then y="./"
			TemplateTagsDic.Item("ZC_BLOG_HOST")=y
		End If


		aryTemplateTagsName=TemplateTagsDic.Keys
		aryTemplateTagsValue=TemplateTagsDic.Items

		Call Filter_Plugin_TArticle_Build_TemplateTags(aryTemplateTagsName,aryTemplateTagsValue)

		Dim s,t

		j=UBound(aryTemplateTagsName)
		For i=1 to j
			If (InStr(aryTemplateTagsName(i),"CACHE_INCLUDE_")>0) And (Right(aryTemplateTagsName(i),5)<>"_HTML") And (Right(aryTemplateTagsName(i),3)<>"_JS") Then
				s=s & aryTemplateTagsName(i) & "|"
			End If
			If ("<#" & aryTemplateTagsName(i) & "#>"="<#CACHE_INCLUDE_CALENDAR#>") Or ("<#" & aryTemplateTagsName(i) & "#>"="<#CACHE_INCLUDE_CALENDAR_JS#>") Or ("<#" & aryTemplateTagsName(i) & "#>"="<#CACHE_INCLUDE_CALENDAR_NOW#>") Then
				aryTemplateTagsValue(i)=Template_Calendar
			End If
		Next


		If IsDynamicLoadSildbar=True Then
			For Each t In Split(s,"|")
				If t="" Then Exit For
				If t<>"CACHE_INCLUDE_NAVBAR" Then
					html=Replace(html,"<#"&t&"#>","<#"&t&"_JS#>")
				End If
			Next
		End If


		j=UBound(aryTemplateTagsName)

		For i=1 to j
			html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
		Next
		html=Replace(html,"<#" & aryTemplateTagsName(0) & "#>",aryTemplateTagsValue(0))

		Call Filter_Plugin_TArticle_Build_Template_Succeed(html)

		Build=True

	End Function


	Public Function Statistic()

		Call GetUsersbyUserIDList(AuthorID)

		'重新统计分类及用户的文章数、评论数
		Dim strSQL,ary1(2),ary2(2),ary3(2)
		strSQL="SELECT"
		ary1(0)=True
		ary2(0)="(SELECT COUNT([log_ID]) FROM [blog_Comment] WHERE [log_ID] =" & ID & " AND [comm_ParentID]=0 AND [comm_isCheck]=0) As [comm_count]"
		ary3(0)="SELECT COUNT([log_ID]) As [comm_count] FROM [blog_Comment] WHERE [log_ID] =" & ID & " AND [comm_ParentID]=0 AND [comm_isCheck]=0"


		ary1(1)=IIf(Categorys(CateID).ReCount=0,True,False)
		ary2(1)="(SELECT COUNT([log_ID]) FROM [blog_Article] WHERE [log_Level]>1 AND [log_Type]=0 AND [log_CateID]=" & CateID & ") As [cate_count]"
		ary3(1)="SELECT COUNT([log_ID]) As [cate_count] FROM [blog_Article] WHERE [log_Level]>1 AND [log_Type]=0 AND [log_CateID]=" & CateID


		ary1(2)=IIf(Users(AuthorID).ReCount=0,True,False)
		ary2(2)="(SELECT COUNT([log_ID]) FROM [blog_Article] WHERE [log_Level]>1 AND [log_Type]=0 AND [log_AuthorID]=" & AuthorID &") As [author_count]"
		ary3(2)="SELECT COUNT([log_ID]) As [author_count] FROM [blog_Article] WHERE [log_Level]>1 AND [log_Type]=0 AND [log_AuthorID]=" & AuthorID

		strSQL=strSQL & IIf(ary1(0),ary2(0),"")
		strSQL=strSQL & IIf(ary1(1),IIf(strSQL=""," ",",") & ary2(1),"")
		strSQL=strSQL & IIf(ary1(2),IIf(strSQL=""," ",",") & ary2(2),"")
		Dim objRS,i,isBool
		i=0
		If ZC_MSSQL_ENABLE Then
			Set objRS=objConn.Execute(strSQL)
			isBool=(Not objRS.bof) And (Not objRS.eof)
		Else
			isBool=True
		End If

		If isBool Then

			If ary1(0) Then

				If Not ZC_MSSQL_ENABLE Then Set objRS=objConn.Execute(ary3(0))
				CommNums=objRs("comm_count")
				objConn.Execute("UPDATE [blog_Article] SET [log_CommNums]="& CommNums &" WHERE [log_ID] =" & ID)

			End If

			If ary1(1) Then

				If Not ZC_MSSQL_ENABLE Then Set objRS=objConn.Execute(ary3(1))
				Call BlogConfig.Write("ZC_UNCATEGORIZED_COUNT",1)
				Categorys(CateID).ReCount=objRs("cate_count")
				If CateID=0 Then
					Call BlogConfig.Write("ZC_UNCATEGORIZED_COUNT",Categorys(CateID).ReCount)
				Else
					objConn.Execute("UPDATE [blog_Category] SET [cate_Count]="&Categorys(CateID).ReCount&" WHERE [cate_ID] =" & CateID)
				End If
				Categorys(CateID).Count=Categorys(CateID).ReCount
				'原分类计数－1
				Dim Cate_ExID:Cate_ExID=Categorys(CateID).ExID
				If Cate_ExID<>-1 Then
					Categorys(Cate_ExID).ReCount=Categorys(Cate_ExID).Count-1
					If Cate_ExID=0 Then
						Call BlogConfig.Write("ZC_UNCATEGORIZED_COUNT",Categorys(Cate_ExID).ReCount)
					Else
						objConn.Execute("UPDATE [blog_Category] SET [cate_Count]="&Categorys(Cate_ExID).ReCount&" WHERE [cate_ID] =" & Cate_ExID)
					End If
					Categorys(Cate_ExID).Count=Categorys(Cate_ExID).ReCount
				End If
				BlogConfig.Save

			End If

			If ary1(2) Then

				If Not ZC_MSSQL_ENABLE Then Set objRS=objConn.Execute(ary3(2))
				Users(AuthorID).ReCount=objRs("author_count")
				objConn.Execute("UPDATE [blog_Member] SET [mem_PostLogs]="&Users(AuthorID).ReCount&" WHERE [mem_ID] =" & AuthorID)
				Users(AuthorID).Count=Users(AuthorID).ReCount
				'原用户计数－1
				Dim User_ExID:User_ExID=Users(AuthorID).ExID
				If User_ExID<>-1 Then

					Users(User_ExID).ReCount=Users(User_ExID).Count-1
					objConn.Execute("UPDATE [blog_Member] SET [mem_PostLogs]="&Users(User_ExID).ReCount&" WHERE [mem_ID] =" & User_ExID)
					Users(User_ExID).Count=Users(User_ExID).ReCount

				End If

			End If

		End If

		objRs.Close
		Set objRs=Nothing

		'FullUrl=Replace(Url,BlogHost,"<#ZC_BLOG_HOST#>")
		'objConn.Execute("UPDATE [blog_Article] SET [log_FullUrl]='"&FullUrl&"' WHERE [log_ID] =" & ID)

		Statistic=True

	End Function


	Function SetVar(TemplateTag,TemplateValue)

		If IsEmpty(html) Then html=Template

		html=Replace(html,"<#" & TemplateTag & "#>",TemplateValue)

	End Function


	Function Save()

		'plugin node
		bAction_Plugin_TArticle_Save_Begin=False
		For Each sAction_Plugin_TArticle_Save_Begin in Action_Plugin_TArticle_Save_Begin
			If Not IsEmpty(sAction_Plugin_TArticle_Save_Begin) Then Call Execute(sAction_Plugin_TArticle_Save_Begin)
			If bAction_Plugin_TArticle_Save_Begin=True Then Exit Function
		Next

		If ZC_POST_STATIC_MODE<>"STATIC" Then Exit Function

		If Not(Level>2) Then Save=True:Exit Function

		Dim objStream

		html=TransferHTML(html,"[no-asp]")

		If ZC_STATIC_TYPE="asp" Then
			html="<"&"%@ CODEPAGE=65001 %"&">" & html
		End If

		Call CreatDirectoryByCustomDirectoryWithFullBlogPath(FullPath)

		Call SaveToFile(FullPath,html,"utf-8",False)

		Save=True

	End Function


	Public Function DelFile()

		Call DelToFile(FullPath)

	End Function


	Public Function Del()

		Call Filter_Plugin_TArticle_Del(ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,FullUrl,FType,MetaString)

		Call DelFile()

		Call CheckParameter(ID,"int",0)
		If (ID=0) Then Del=False:Exit Function

		objConn.Execute("DELETE FROM [blog_Article] WHERE [log_ID] =" & ID)
		objConn.Execute("DELETE FROM [blog_Comment] WHERE [log_ID] =" & ID)
		objConn.Execute("DELETE FROM [blog_TrackBack] WHERE [log_ID] =" & ID)

		Del=True

	End Function



	Function SaveCache()
	End Function


	Function LoadCache()
	End Function


	Private Sub Class_Initialize()
		PostTime=GetTime(Now())
		PostTime=Year(PostTime) & "-" & Month(PostTime) & "-" & Day(PostTime) & " " & Hour(PostTime) & ":" & Minute(PostTime) & ":" & Second(PostTime)
		ID=0
		CateID=0
		AuthorID=0
		Level=4'默认为普通
		Title=ZC_MSG099
		IP=GetReallyIP

		IsDynamicLoadSildbar=True

		Ftemplate=Empty
		FType=ZC_POST_TYPE_ARTICLE

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

	Public Template_PageBar
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
	Public subhtml
	Public subhtml_TemplateName

	Public IsDynamicLoadSildbar
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

		'plugin node
		bAction_Plugin_TArticleList_Export_Begin=False
		For Each sAction_Plugin_TArticleList_Export_Begin in Action_Plugin_TArticleList_Export_Begin
			If Not IsEmpty(sAction_Plugin_TArticleList_Export_Begin) Then Call Execute(sAction_Plugin_TArticleList_Export_Begin)
			If bAction_Plugin_TArticleList_Export_Begin=True Then Exit Function
		Next

		Call Filter_Plugin_TArticleList_Export(intPage,anyCate,anyAuthor,dtmDate,anyTag,intType)

		ListType="DEFAULT"
		Url =ParseCustomDirectoryForUrl(FullRegex,ZC_STATIC_DIRECTORY,"","","","","","","","")
		MixUrl=ParseCustomDirectoryForUrl("{%host%}/catalog.asp",ZC_STATIC_DIRECTORY,"","","","","","","","")

		Dim aryArticle
		Dim aryArticleList()
		Dim  Template_Article_Istop


		'plugin node
		Dim i,j,k,l,s,t
		Dim objRS
		Dim intPageCount
		Dim objArticle
		Dim ut,ud,dd,dt,tt,td
		Dim intCate,intAuthor,intTag

		Call CheckParameter(intPage,"int",1)
		Title=ZC_BLOG_SUBTITLE

		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn

		Set dt = CreateObject("Scripting.Dictionary")
		Set dd = CreateObject("Scripting.Dictionary")

		'//////////////////////////
		'ontop
		objRS.Source="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) AND ([log_Level]>1) AND ([log_Istop]<>0)"
		objRS.Source=objRS.Source & "ORDER BY [log_PostTime] DESC"
		objRS.Open()
		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = ZC_DISPLAY_COUNT
			intPageCount=objRS.PageCount
			objRS.AbsolutePage = 1

			For i = 1 To objRS.PageSize

				Set objArticle=New TArticle
				If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then
					ut=ut & "," & objRs(7)
					tt=tt & objRs(1)
					dt.Add CLng(objRs(0)), objArticle
				End If
				Set objArticle=Nothing

				objRS.MoveNext
				If objRS.EOF Then Exit For

			Next

		End If
		objRS.Close()
		'//////////////////////////


		objRS.Source="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) AND ([log_Level]>1)"

		If anyAuthor<>"" Then

			ListType="USER"

			If InStr(ZC_USER_REGEX,"{%alias%}")>0 Then
				i=GetAuthorByAlias(anyAuthor)
				If i=-1 Then
					i=GetAuthorByName(anyAuthor)
				End If
			ElseIf InStr(ZC_USER_REGEX,"{%name%}")>0 Then
				i=GetAuthorByName(Author)
			Else
				i=anyAuthor
			End If

			Call CheckParameter(i,"int",Empty)

			intAuthor=i

			objRS.Source=objRS.Source & "AND([log_AuthorID]="&i&")"

			If CheckAuthorByID(i) Then
				Call GetUsersbyUserIDList(i)
				Title=Users(i).Name
				TemplateTags_ArticleList_Author_ID=Users(i).ID
				If IsEmpty(html)=True Then html=Users(i).Template
				Url =ParseCustomDirectoryForUrl(Users(i).FullRegex,ZC_STATIC_DIRECTORY,"","","","","",Users(i).ID,Users(i).Name,Users(i).StaticName)
				MixUrl=ParseCustomDirectoryForUrl("{%host%}/catalog.asp?auth={%id%}",ZC_STATIC_DIRECTORY,"","","","","",Users(i).ID,Users(i).Name,Users(i).StaticName)
			End If
		End if
		If anyCate<>"" Then

			ListType="CATEGORY"

			If InStr(ZC_CATEGORY_REGEX,"{%alias%}")>0 Then
				i=GetCateByAlias(anyCate)
				If i=-1 Then
					i=GetCateByName(anyCate)
				End If
			ElseIf InStr(ZC_CATEGORY_REGEX,"{%name%}")>0 Then
				i=GetCateByName(anyCate)
			Else
				i=anyCate
			End If

			Call CheckParameter(i,"int",Empty)

			intCate=i

			If i=0 Then
				objRS.Source=objRS.Source & "AND([log_CateID]=0)"
			Else
				Dim strSubCateID
				strSubCateID=Join(GetSubCateID(i,True),",")
				objRS.Source=objRS.Source & "AND([log_CateID]IN("&strSubCateID&"))"
			End If

			If CheckCateByID(i) Then
				Title=Categorys(i).Name
				TemplateTags_ArticleList_Category_ID=Categorys(i).ID
				If IsEmpty(html)=True Then html=Categorys(i).Template
				Url =ParseCustomDirectoryForUrl(Categorys(i).FullRegex,ZC_STATIC_DIRECTORY,"","","","","",Categorys(i).ID,Categorys(i).Name,Categorys(i).StaticName)
				MixUrl =ParseCustomDirectoryForUrl("{%host%}/catalog.asp?cate={%id%}",ZC_STATIC_DIRECTORY,"","","","","",Categorys(i).ID,Categorys(i).Name,Categorys(i).StaticName)
			End If
		End if
		If IsDate(dtmDate) Then

			ListType="DATE"

			Dim y
			Dim m
			Dim d
			Dim ny
			Dim nm

			y=Year(dtmDate)
			m=Month(dtmDate)
			d=Day(dtmDate)

			If Not Len(CStr(dtmDate))>7 Then
				d=""
			End If

			Url =UrlbyDate(y,m,d)
			MixUrl =UrlbyDateAuto(y,m,d)

			TemplateTags_ArticleList_Date_ShortDate=dtmDate
			TemplateTags_ArticleList_Date_Year=y
			TemplateTags_ArticleList_Date_Month=m
			TemplateTags_ArticleList_Date_Day=d

			ny=y
			nm=m+1
			If m=12 Then ny=ny+1:nm=1

			If Len(CStr(dtmDate))>7 Then
				objRS.Source=objRS.Source & "AND(Year([log_PostTime])="&y&") AND(Month([log_PostTime])="&m&") AND(Day([log_PostTime])="&d&")"
			Else
				objRS.Source=objRS.Source & "AND(Year([log_PostTime])="&y&") AND(Month([log_PostTime])="&m&")"
			End If

			Template_Calendar="<script src=""<#ZC_BLOG_HOST#>zb_system/function/c_html_js.asp?date="&dtmDate&""" type=""text/javascript""></script>"

			If IsEmpty(html)=True Then html=GetTemplate("TEMPLATE_CATALOG")

			Title=Year(dtmDate) & " " & ZVA_Month(Month(dtmDate)) & IIF(Len(CStr(dtmDate))>7," " & Day(dtmDate),"")
		End If
		If anyTag<>"" Then

			ListType="TAGS"


			If InStr(ZC_TAGS_REGEX,"{%alias%}")>0 Then
				If CheckTagByIntro(anyTag) Then
					i=GetTagByIntro(anyTag)
				Else
					If CheckTagByName(anyTag) Then i=GetTagByName(anyTag)
				End If
			ElseIf InStr(ZC_TAGS_REGEX,"{%id%}")>0 Then
				i=CLng(anyTag)
			Else
				If CheckTagByName(anyTag) Then i=GetTagByName(anyTag)
			End If

			objRS.Source=objRS.Source & "AND([log_Tag] LIKE '%{" & i & "}%')"

			intTag=i

			If CheckTagByID(i) Then
				Call GetTagsbyTagIDList("{"&i&"}")

				Title=Tags(i).Name
				TemplateTags_ArticleList_Tags_ID=Tags(i).ID
				If IsEmpty(html)=True Then html=Tags(i).Template
				Url =ParseCustomDirectoryForUrl(Tags(i).FullRegex,ZC_STATIC_DIRECTORY,"","","","","",Tags(i).ID,Tags(i).Name,Tags(i).EncodeName)
				MixUrl =ParseCustomDirectoryForUrl("{%host%}/catalog.asp?tags={%alias%}",ZC_STATIC_DIRECTORY,"","","","","",Tags(i).ID,Tags(i).Name,Tags(i).EncodeName)
			End If
		End If


		If ListType="DEFAULT" Then objRS.Source=objRS.Source & " AND ([log_Istop]=0) "


		objRS.Source=objRS.Source & "ORDER BY [log_PostTime] DESC"
		objRS.Open()

		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = ZC_DISPLAY_COUNT
			intPageCount=objRS.PageCount
			objRS.AbsolutePage = intPage

			For i = 1 To objRS.PageSize

				If intPage>intPageCount Then Exit For

				Set objArticle=New TArticle
				If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then
					ud=ud & "," & objRs(7)
					td=td & objRs(1)
					dd.Add CLng(objRs(0)), objArticle
				End If
				Set objArticle=Nothing

				objRS.MoveNext
				If objRS.EOF Then Exit For

			Next

		End If

		objRS.Close()
		Set objRS=Nothing


		If IsEmpty(html)=True Then
			If intPage>1 And ZC_DEFAULT_PAGES_TEMPLATE<>"" Then
				html=GetTemplate("TEMPLATE_" & ZC_DEFAULT_PAGES_TEMPLATE)
			Else
				html=Template
			End If
		End If

		Call GetTagsbyTagIDList(tt & td)

		Call GetUsersbyUserIDList(ut & "," & ud)



		ReDim aryArticleList(-1)
		For i=0 To dt.Count-1
			j=dt.Keys
			ReDim Preserve aryArticleList(i)
			Set objArticle=dt.Item(j(i))
			objArticle.html=html
			If objArticle.Export(ZC_DISPLAY_MODE_ONTOP)= True Then
				aryArticleList(i)=objArticle.subhtml
			End If
		Next
		Template_Article_Istop=Join(aryArticleList)


		Dim RE ,Match,Matches
		ReDim aryArticleList(-1)
		If dd.Count>0 THen
			For i=0 To dd.Count-1
				j=dd.Keys
				ReDim Preserve aryArticleList(i)
				Set objArticle=dd.Item(j(i))
				objArticle.html=html
				If objArticle.Export(intType)= True Then
					aryArticleList(i)=objArticle.subhtml
				End If
				subhtml_TemplateName=objArticle.subhtml_TemplateName
			Next
		Else
			Set RE = New RegExp
				RE.Pattern = "\<\#template\:(article\-([a-z0-9]*)([\-a-z0-9]*))\#\>"
				RE.IgnoreCase = True
				RE.Global = True
				Set Matches = RE.Execute(html)
				For Each Match in Matches
					If IsEmpty(subhtml_TemplateName) Then subhtml_TemplateName="template:"&Match.SubMatches(0)&""
					Exit For
				Next
				Set Matches = Nothing
			Set RE = Nothing
		End If
		subhtml=Join(aryArticleList)


		If ListType="DEFAULT" Then subhtml=Template_Article_Istop & subhtml


		TemplateTags_ArticleList_Page_Now=intPage
		TemplateTags_ArticleList_Page_All=intPageCount

    If dt.Count=0 And dd.Count = 0 Then
      Export = False
      Exit Function
    End If
    If ListType <> "DEFAULT" And dd.Count = 0 Then
      Export = False
      Exit Function
    End If
    If intPageCount<intPage Then
      Export = False
      Exit Function
    End If

		Call ExportBar(intPage,intPageCount)


		Dim aryTemplateSubName()
		Dim aryTemplateSubValue()


		ReDim aryTemplateSubName( 47)
		ReDim aryTemplateSubValue(47)

		aryTemplateSubName(  1)=subhtml_TemplateName
		aryTemplateSubValue( 1)=subhtml
		aryTemplateSubName(  2)="template:pagebar"
		aryTemplateSubValue( 2)=TransferHTML(Template_PageBar,"[anti-zc_blog_host]")
		aryTemplateSubName(  3)="template:pagebar_next"
		aryTemplateSubValue( 3)=TransferHTML(Template_PageBar_Next,"[anti-zc_blog_host]")
		aryTemplateSubName(  4)="template:pagebar_previous"
		aryTemplateSubValue( 4)=TransferHTML(Template_PageBar_Previous,"[anti-zc_blog_host]")

		aryTemplateSubName(  6)="articlelist/date/year"
		aryTemplateSubValue( 6)=TemplateTags_ArticleList_Date_Year
		aryTemplateSubName(  7)="articlelist/date/month"
		aryTemplateSubValue( 7)=TemplateTags_ArticleList_Date_Month
		aryTemplateSubName(  9)="articlelist/date/day"
		aryTemplateSubValue( 9)=TemplateTags_ArticleList_Date_Day
		aryTemplateSubName( 10)="articlelist/date/shortdate"
		aryTemplateSubValue(10)=TemplateTags_ArticleList_Date_ShortDate
		aryTemplateSubName( 11)="articlelist/page/now"
		aryTemplateSubValue(11)=TemplateTags_ArticleList_Page_Now
		aryTemplateSubName( 12)="articlelist/page/all"
		aryTemplateSubValue(12)=TemplateTags_ArticleList_Page_All
		aryTemplateSubName( 13)="articlelist/page/count"
		aryTemplateSubValue(13)=ZC_DISPLAY_COUNT



			aryTemplateSubName( 14)="articlelist/category/id"
			aryTemplateSubName( 15)="articlelist/category/name"
			aryTemplateSubName( 16)="articlelist/category/order"
			aryTemplateSubName( 17)="articlelist/category/count"
			aryTemplateSubName( 18)="articlelist/category/url"
			aryTemplateSubName( 19)="articlelist/category/staticname"

			aryTemplateSubName( 20)="articlelist/category/parent/id"
			aryTemplateSubName( 21)="articlelist/category/parent/name"
			aryTemplateSubName( 22)="articlelist/category/parent/order"
			aryTemplateSubName( 23)="articlelist/category/parent/count"
			aryTemplateSubName( 24)="articlelist/category/parent/url"
			aryTemplateSubName( 25)="articlelist/category/parent/staticname"


		If ListType="CATEGORY" Then
			aryTemplateSubValue(14)=Categorys(intCate).ID
			aryTemplateSubValue(15)=Categorys(intCate).HtmlName
			aryTemplateSubValue(16)=Categorys(intCate).Order
			aryTemplateSubValue(17)=Categorys(intCate).Count
			aryTemplateSubValue(18)=TransferHTML(Categorys(intCate).HtmlUrl,"[anti-zc_blog_host]")
			aryTemplateSubValue(19)=Categorys(intCate).StaticName

			If Categorys(intCate).ParentID<>0 Then
			aryTemplateSubValue(20)=Categorys(Categorys(intCate).ParentID).ID
			aryTemplateSubValue(21)=Categorys(Categorys(intCate).ParentID).HtmlName
			aryTemplateSubValue(22)=Categorys(Categorys(intCate).ParentID).Order
			aryTemplateSubValue(23)=Categorys(Categorys(intCate).ParentID).Count
			aryTemplateSubValue(24)=TransferHTML(Categorys(Categorys(intCate).ParentID).HtmlUrl,"[anti-zc_blog_host]")
			aryTemplateSubValue(25)=Categorys(Categorys(intCate).ParentID).StaticName
			End If
		End If

			aryTemplateSubName( 26)="articlelist/author/id"
			aryTemplateSubName( 27)="articlelist/author/name"
			aryTemplateSubName( 28)="articlelist/author/level"
			aryTemplateSubName( 29)="articlelist/author/email"
			aryTemplateSubName( 30)="articlelist/author/homepage"
			aryTemplateSubName( 31)="articlelist/author/count"
			aryTemplateSubName( 32)="articlelist/author/url"
			aryTemplateSubName( 33)="articlelist/author/staticname"
			aryTemplateSubName( 34)="articlelist/author/intro"
			aryTemplateSubName( 35)="articlelist/author/avatar"
			aryTemplateSubName( 36)="articlelist/author/levelname"

		If ListType="USER" Then
			aryTemplateSubValue(26)=Users(intAuthor).ID
			aryTemplateSubValue(27)=Users(intAuthor).FirstName
			aryTemplateSubValue(28)=Users(intAuthor).Level
			aryTemplateSubValue(29)=Users(intAuthor).Email
			aryTemplateSubValue(30)=Users(intAuthor).HomePage
			aryTemplateSubValue(31)=Users(intAuthor).Count
			aryTemplateSubValue(32)=TransferHTML(Users(intAuthor).HtmlUrl,"[anti-zc_blog_host]")
			aryTemplateSubValue(33)=Users(intAuthor).StaticName
			aryTemplateSubValue(34)=Users(intAuthor).Intro
			aryTemplateSubValue(35)=Users(intAuthor).Avatar
			aryTemplateSubValue(36)=Users(intAuthor).LevelName
		End If

			aryTemplateSubName( 37)="articlelist/tag/id"
			aryTemplateSubName( 38)="articlelist/tag/name"
			aryTemplateSubName( 39)="articlelist/tag/intro"
			aryTemplateSubName( 40)="articlelist/tag/count"
			aryTemplateSubName( 41)="articlelist/tag/url"
			aryTemplateSubName( 42)="articlelist/tag/encodename"

		If ListType="TAGS" Then
			aryTemplateSubValue(37)=Tags(intTag).ID
			aryTemplateSubValue(38)=Tags(intTag).HtmlName
			aryTemplateSubValue(39)=Tags(intTag).HtmlIntro
			aryTemplateSubValue(40)=Tags(intTag).Count
			aryTemplateSubValue(41)=Tags(intTag).HtmlUrl
			aryTemplateSubValue(42)=Tags(intTag).EncodeName
		End If

		aryTemplateSubName( 43)="template:sidebar"
		aryTemplateSubValue(43)=GetTemplate("CACHE_SIDEBAR")
		aryTemplateSubName( 44)="template:sidebar2"
		aryTemplateSubValue(44)=GetTemplate("CACHE_SIDEBAR2")
		aryTemplateSubName( 45)="template:sidebar3"
		aryTemplateSubValue(45)=GetTemplate("CACHE_SIDEBAR3")
		aryTemplateSubName( 46)="template:sidebar4"
		aryTemplateSubValue(46)=GetTemplate("CACHE_SIDEBAR4")
		aryTemplateSubName( 47)="template:sidebar5"
		aryTemplateSubValue(47)=GetTemplate("CACHE_SIDEBAR5")


		'plugin node
		Call Filter_Plugin_TArticleList_Export_TemplateTags(aryTemplateSubName,aryTemplateSubValue)

		j=UBound(aryTemplateSubName)
		For i=0 to j
			If IsNull(aryTemplateSubValue(i))=True Then aryTemplateSubValue(i)=""
			html=Replace(html,"<#" & aryTemplateSubName(i) & "#>",aryTemplateSubValue(i))
		Next



		Dim da
		Set da=CreateObject("Scripting.Dictionary")

		Set RE = New RegExp
		RE.IgnoreCase = True
		RE.Global = True

		If Categorys(intCate).ParentID=0 Then
			RE.Pattern = "<#template:articlelist_category_parent:begin#>(.|\n)*<#template:articlelist_category_parent:end#>"
			html = RE.Replace(html, "")
		Else
			html=Replace(html,"<#template:articlelist_category_parent:begin#>","")
			html=Replace(html,"<#template:articlelist_category_parent:end#>","")
		End If
		Set RE = Nothing


		Set RE = New RegExp
		RE.Pattern = "<#articlelist/((category/meta|author/meta)/([a-z0-9_]{1,}))#>"
		RE.IgnoreCase = True
		RE.Global = True
		Set Matches = RE.Execute(html)
		For Each Match in Matches
			s=Match.SubMatches(0)
			If da.Exists(s)=False Then da.add s,s
		Next
		Set Matches = Nothing
		Set RE = Nothing

		s=da.Items
		For t = 0 To da.Count -1
			If InStr(s(t),"category/meta/")>0 Then
				s(t)=Replace(s(t),"category/meta/","")
				If ListType="CATEGORY" Then
					html = Replace(html,"<#articlelist/category/meta/" & s(t) & "#>",Categorys(intCate).Meta.GetValue(s(t)) )
				Else
					html = Replace(html,"<#articlelist/category/meta/" & s(t) & "#>","")
				End If
			ElseIf InStr(s(t),"author/meta/")>0 Then
				s(t)=Replace(s(t),"author/meta/","")
				If ListType="USER" Then
					html = Replace(html,"<#articlelist/author/meta/" & s(t) & "#>",Users(intAuthor).Meta.GetValue(s(t)) )
				Else
					html = Replace(html,"<#articlelist/author/meta/" & s(t) & "#>","")
				End If
			End If
		Next

		Url=Replace(Replace(Url,"//","/"),":/","://",1,1)

		Export=True

		'plugin node
		bAction_Plugin_TArticleList_Export_End=False
		For Each sAction_Plugin_TArticleList_Export_End in Action_Plugin_TArticleList_Export_End
			If Not IsEmpty(sAction_Plugin_TArticleList_Export_End) Then Call Execute(sAction_Plugin_TArticleList_Export_End)
			If bAction_Plugin_TArticleList_Export_End=True Then Exit Function
		Next

	End Function



	Public Function Build()

		Dim i,j

		'plugin node
		Call Filter_Plugin_TArticleList_Build_Template(html)

		Dim aryTemplateTagsName
		Dim aryTemplateTagsValue

		TemplateTagsDic.Item("BlogTitle")=HtmlTitle
		TemplateTagsDic.Item("ZC_BLOG_HOST")=BlogHost

		If Template_Calendar="" Then Template_Calendar="<script src=""<#ZC_BLOG_HOST#>zb_system/function/c_html_js.asp?date=now"" type=""text/javascript""></script>"

		If ZC_MULTI_DOMAIN_SUPPORT=True And ZC_PERMANENT_DOMAIN_ENABLE=False Then
			Dim x,y
			x=CStr(Replace(Url,BlogHost,""))
			x=Replace(x,"/////","/"):x=Replace(x,"////","/"):x=Replace(x,"///","/"):x=Replace(x,"//","/")
			For i=1 To UBound(Split(x,"/"))
				y=y & "../"
			Next
			If y="" Then y="./"
			TemplateTagsDic.Item("ZC_BLOG_HOST")=y
		End If

		aryTemplateTagsName=TemplateTagsDic.Keys
		aryTemplateTagsValue=TemplateTagsDic.Items

		Call Filter_Plugin_TArticleList_Build_TemplateTags(aryTemplateTagsName,aryTemplateTagsValue)

		Dim s,t
		j=UBound(aryTemplateTagsName)
		For i=1 to j
			If (InStr(aryTemplateTagsName(i),"CACHE_INCLUDE_")>0) And (Right(aryTemplateTagsName(i),5)<>"_HTML") And (Right(aryTemplateTagsName(i),3)<>"_JS") Then
				s=s & aryTemplateTagsName(i) & "|"
			End If
			If ("<#" & aryTemplateTagsName(i) & "#>"="<#CACHE_INCLUDE_CALENDAR#>") Or ("<#" & aryTemplateTagsName(i) & "#>"="<#CACHE_INCLUDE_CALENDAR_JS#>") Or ("<#" & aryTemplateTagsName(i) & "#>"="<#CACHE_INCLUDE_CALENDAR_NOW#>") Then
				aryTemplateTagsValue(i)=Template_Calendar
			End If
		Next

		If IsDynamicLoadSildbar=True Then
			For Each t In Split(s,"|")
				If t="" Then Exit For
				If t<>"CACHE_INCLUDE_NAVBAR" Then
					html=Replace(html,"<#"&t&"#>","<#"&t&"_JS#>")
				End If
			Next
		End If

		j=UBound(aryTemplateTagsName)
		For i=1 to j
			html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
		Next
		html=Replace(html,"<#" & aryTemplateTagsName(0) & "#>",aryTemplateTagsValue(0))

		'plugin node
		Call Filter_Plugin_TArticleList_Build_Template_Succeed(html)

		Build=True

	End Function


	Public Function ExportBar(intNowPage,intAllPage)

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

		t=Url

		Dim QuerySplit
		QuerySplit=IIf(InStr(t,"?")>0,"&","?")

		'If ListType="CATEGORY" Or ListType="USER" Or ListType="DATE" Or ListType="TAGS" Then
			If ZC_STATIC_MODE="ACTIVE" Then
				t=t & QuerySplit & "page=%n"
			End If
			If ZC_STATIC_MODE="REWRITE" Then
				If InStr(t,"{%page%}")>0 Then
					t=Replace(t,"default.html","")
					t=Replace(t,"{%page%}","%n")
				ElseIf InStr(t,"/default.html")>0 Then
					If InStr(ZC_DEFAULT_REGEX,"{%page%}")>0 Then
						t=Replace(t,"/default.html","/%n/default.html")
					ElseIf ListType="DEFAULT" Then
						t=Replace(t,"/default.html","/default_%n.html")
					Else
						t=Replace(t,"/default.html","_%n/default.html")
					End If
				Else
					t=Replace(t,".html","_%n.html")
				End If
			End If
			If ZC_STATIC_MODE="MIX" Then
				t=MixUrl
				t=t & QuerySplit & "page=%n"
			End If
		'End If

		If intAllPage>0 Then
			Dim a,b

			s=Replace(t,"%n",1)
			If ListType="DEFAULT" Then s=BlogHost
			If (ListType="CATEGORY" Or ListType="USER" Or ListType="DATE" Or ListType="TAGS") Then s=Url
			If ZC_STATIC_MODE="REWRITE" Then s=Replace(s,"/default.html","/")

			strPageBar=GetTemplate("TEMPLATE_B_PAGEBAR")
			strPageBar=Replace(strPageBar,"<#pagebar/page/url#>",s)
			strPageBar=Replace(strPageBar,"<#pagebar/page/number#>","<span class=""page first-page"">"&ZC_MSG235&"</span>")
			Template_PageBar=Template_PageBar & strPageBar

			If intAllPage>ZC_PAGEBAR_COUNT Then
				a=intNowPage
				b=intNowPage+ZC_PAGEBAR_COUNT-1
				If a>ZC_PAGEBAR_COUNT Then a=a-1:b=b-1
				If b>intAllPage Then b=intAllPage:a=intAllPage-ZC_PAGEBAR_COUNT+1
			Else
				a=1:b=intAllPage
			End If
			For i=a to b

				s=Replace(t,"%n",i)
				If ListType="DEFAULT" And i=1 Then s=BlogHost
				If (ListType="CATEGORY" Or ListType="USER" Or ListType="DATE" Or ListType="TAGS") And i=1 Then s=Url
				If ZC_STATIC_MODE="REWRITE" Then s=Replace(s,"/default.html","/")

				strPageBar=GetTemplate("TEMPLATE_B_PAGEBAR")
				If i=intNowPage then
					Template_PageBar=Template_PageBar & "<span class=""page now-page"">" & i & "</span>"
				Else
					strPageBar=Replace(strPageBar,"<#pagebar/page/url#>",s)
					strPageBar=Replace(strPageBar,"<#pagebar/page/number#>","<span class=""page"">"&i&"</span>")
					Template_PageBar=Template_PageBar & strPageBar
				End If

			Next

			s=Replace(t,"%n",intAllPage)
			If ListType="DEFAULT" And intAllPage=1 Then s=BlogHost
			If (ListType="CATEGORY" Or ListType="USER" Or ListType="DATE" Or ListType="TAGS") And intAllPage=1 Then s=Url
			If ZC_STATIC_MODE="REWRITE" Then s=Replace(s,"/default.html","/")

			strPageBar=GetTemplate("TEMPLATE_B_PAGEBAR")
			strPageBar=Replace(strPageBar,"<#pagebar/page/url#>",s)
			strPageBar=Replace(strPageBar,"<#pagebar/page/number#>","<span class=""page last-page"">"&ZC_MSG236&"</span>")
			Template_PageBar=Template_PageBar & strPageBar

			If intNowPage=1 Then
				Template_PageBar_Previous=""
			Else
				s=Replace(t,"%n",intNowPage-1)

				If ListType="DEFAULT" And intNowPage-1=1 Then s=BlogHost
				If (ListType="CATEGORY" Or ListType="USER" Or ListType="DATE" Or ListType="TAGS") And intNowPage-1=1 Then s=Url
				If ZC_STATIC_MODE="REWRITE" Then s=Replace(s,"/default.html","/")

				Template_PageBar_Previous="<span class=""pagebar-previous""><a href="""& s &"""><span>"&ZC_MSG156&"</span></a></span>"

			End If

			If intNowPage=intAllPage Then
				Template_PageBar_Next=""
			Else
				s=Replace(t,"%n",intNowPage+1)
				If ZC_STATIC_MODE="REWRITE" Then s=Replace(s,"/default.html","/")
				Template_PageBar_Next="<span class=""pagebar-next""><a href="""& s &"""><span>"&ZC_MSG155&"</span></a></span>"
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


	Function SetVar(TemplateTag,TemplateValue)

		If IsEmpty(html) Then html=Template

		html=Replace(html,"<#" & TemplateTag & "#>",TemplateValue)

	End Function


	Function Save()

		html=TransferHTML(html,"[no-asp]")
		If ZC_STATIC_TYPE="asp" Then
			html="<"&"%@ CODEPAGE=65001 %"&">" & html
		End If

		Call CreatDirectoryByCustomDirectoryWithFullBlogPath(FullPath)

		Call SaveToFile(FullPath,html,"utf-8",False)

		Save=True

	End Function


	Function SaveCache()
	End Function

	Function LoadCache()
	End Function


	Private Sub Class_Initialize()

		IsDynamicLoadSildbar=False
		ListType="DEFAULT"'CATEGORY'USER'DATE'TAGS
		'isCatalog=False

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

	Public Property Get LevelName
		LevelName=ZVA_User_Level_Name(Level)
	End Property

	Private Ffullregex
	Public Property Let FullRegex(s)
		Ffullregex=s
	End Property
	Public Property Get FullRegex
		If Ffullregex<>"" Then
			FullRegex=Ffullregex
		Else
			FullRegex=ZC_USER_REGEX
		End If
	End Property

	Public FullUrl
	Public TemplateName


	Private Ftemplate
	Public Property Let Template(strFileName)
		Ftemplate=GetTemplate("TEMPLATE_" & strFileName)
	End Property
	Public Property Get Template
		If Ftemplate<>"" Then
			Template = Ftemplate
			Exit Property
		Else
			If TemplateName<>"" Then
				Dim s
				s=GetTemplate("TEMPLATE_" &TemplateName)
				If s<>"" Then
					Ftemplate = s
				Else
					Ftemplate=GetTemplate("TEMPLATE_CATALOG")
				End If
			Else
				Ftemplate=GetTemplate("TEMPLATE_CATALOG")
			End If
			Template = Ftemplate
		End If
	End Property


	Public Function GetDefaultTemplateName
		If TemplateName<>"" Then
			GetDefaultTemplateName=TemplateName
		Else
			GetDefaultTemplateName="CATALOG"
		End If
	End Function


	Public Property Get FullPath
		FullPath=ParseCustomDirectoryForPath(FullRegex,ZC_STATIC_DIRECTORY,"","","","","",ID,Name,StaticName)
	End Property

	Public Property Get Url

		'plugin node
		bAction_Plugin_TUser_Url=False
		For Each sAction_Plugin_TUser_Url in Action_Plugin_TUser_Url
			If Not IsEmpty(sAction_Plugin_TUser_Url) Then Call Execute(sAction_Plugin_TUser_Url)
			If bAction_Plugin_TUser_Url=True Then Exit Property
		Next

		Url =ParseCustomDirectoryForUrl(FullRegex,ZC_STATIC_DIRECTORY,"","","","","",ID,Name,StaticName)
		If Right(Url,12)="default.html" Then Url=Left(Url,Len(Url)-12)

		Url=Replace(Replace(Url,"//","/"),":/","://",1,1)

		Call Filter_Plugin_TUser_Url(Url)

	End Property

	Public Property Get FirstName
		FirstName=IIf(Level=5,Name,IIf(Alias="",Name,Alias))
	End Property

	Private Function GetAvatar
	  If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")
	  If (PublicObjFSO.FileExists(BlogPath & "zb_users/avatar/"&ID&".png")) Then
		GetAvatar=BlogHost & "zb_users/avatar/"&ID&".png"
	  Else
		GetAvatar=BlogHost & "zb_users/avatar/0.png"
	  End If
	End Function


	Public Property Get Avatar

		Avatar=GetAvatar

	End Property

	Public Property Get HtmlUrl
		HtmlUrl=TransferHTML(Url,"[html-format]")
	End Property

	Public Property Get RssUrl
		RssUrl = BlogHost & "feed.asp?auth=" & ID
	End Property


	Public Property Get StaticName
		If IsNull(Alias) Or IsEmpty(Alias) Or Alias="" Then
			StaticName = Name
		Else
			StaticName = Alias
		End If
	End Property


	Private FEmailMD5
	Public Property Get EmailMD5
		If FEmailMD5="" Then
			If Email="" Or IsNull(Email)=True Then
				FEmailMD5=""
			Else
				FEmailMD5=MD5(Email)
			End If
		End If
		EmailMD5=FEmailMD5
	End Property


	Private FLoginType
	Public Property Let LoginType(strLoginType)
			FLoginType=strLoginType
	End Property
	Public Property Get LoginType
			LoginType = FLoginType
	End Property

	Public ReCount
	Public ExID


	Private FGuid
	Public Property Get Guid
		If FGuid="" Then
			FGuid=RndGuid()
			If ID>0 Then
				FGuid=objConn.Execute("SELECT [mem_Guid] FROM [blog_Member] WHERE [mem_ID]="&ID)(0)
			End If
		End If
		Guid=FGuid
	End Property

	Public Function CreatePasswordByOriginal(OriginaPassword)
		CreatePasswordByOriginal=MD5(MD5(OriginaPassword) & Guid)
	End Function

	Public Function GetPasswordByOriginal(OriginaPassword)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [mem_Guid] FROM [blog_Member] WHERE [mem_Name]='"&Name & "'" )
		If (Not objRS.Bof) And (Not objRS.Eof) Then
			GetPasswordByOriginal=MD5(MD5(OriginaPassword) & objRS("mem_Guid"))
		End If

		objRS.Close
		Set objRS=Nothing

	End Function


	Public Function GetPasswordByMD5(Md5Password)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [mem_Guid] FROM [blog_Member] WHERE [mem_Name]='"&Name & "'" )
		If (Not objRS.Bof) And (Not objRS.Eof) Then
			GetPasswordByMD5=MD5(Md5Password & objRS("mem_Guid"))
		End If

		objRS.Close
		Set objRS=Nothing

	End Function


	Public Function Verify()
		Dim strUserName
		Dim strPassWord

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


		If Session(ZC_BLOG_CLSID & "quicklogin")=MD5(Month(Now) & Day(Now) & ZC_BLOG_CLSID & strUserName & strPassWord) Then
			If LoadInfobyArray(Application(ZC_BLOG_CLSID & "QUICKLOGIN_ID" & Session(ZC_BLOG_CLSID & "quicklogin_id"))) Then
				Verify=True
				Exit Function
			End If
		End If

		'校检
		'If Len(strUserName) >ZC_USERNAME_MAX Then Call ShowError(7)
		'If Len(strPassWord)<>32 Then Call ShowError(55)
		'If Not CheckRegExp(strUserName,"[username]") Then Call ShowError(7)

		strUserName=FilterSQL(strUserName)
		strPassWord=FilterSQL(strPassWord)

		Dim objRS
		Set objRS=objConn.Execute("SELECT * FROM [blog_Member] WHERE [mem_Name]='"&strUserName & "'")
		If (Not objRS.Bof) And (Not objRS.Eof) Then

			If StrComp(strPassWord,objRS("mem_Password"))=0 Then

				Call LoadInfobyID(objRS("mem_ID"))

				Session(ZC_BLOG_CLSID & "quicklogin")=MD5(Month(Now) & Day(Now) & ZC_BLOG_CLSID & strUserName & strPassWord)
				Session(ZC_BLOG_CLSID & "quicklogin_id")=ID
				Application.Lock
				Application(ZC_BLOG_CLSID & "QUICKLOGIN_ID" & ID)=Array(ID,Name,Level,Password,Email,HomePage,Count,Alias,TemplateName,FullUrl,Intro,MetaString)
				Application.UnLock

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
		Set objRS=objConn.Execute("SELECT [mem_ID],[mem_Name],[mem_Level],[mem_Password],[mem_Email],[mem_HomePage],[mem_PostLogs],[mem_Intro],[mem_Template],[mem_FullUrl],[mem_Url],[mem_Meta] FROM [blog_Member] WHERE [mem_ID]=" & user_ID)
		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("mem_ID")
			Name=objRS("mem_Name")
			Level=objRS("mem_Level")
			Password=objRS("mem_Password")
			Email=objRS("mem_Email")
			HomePage=objRS("mem_HomePage")
			Count=objRS("mem_PostLogs")
			Alias=objRS("mem_Url")
			TemplateName=objRS("mem_Template")
			FullUrl=objRS("mem_FullUrl")
			Intro=objRS("mem_Intro")
			MetaString=objRS("mem_Meta")

			If IsNull(Email) Or IsEmpty(Email) Or Len(Email)=0 Then Email="null@null.com"
			If IsNull(HomePage) Then HomePage=""
			If IsNull(Alias) Then Alias=""
			If IsNull(TemplateName) Then TemplateName=""

			LoadInfobyID=True
		End If

		objRS.Close
		Set objRS=Nothing


		Call Filter_Plugin_TUser_LoadInfobyID(ID,Name,Level,Password,Email,HomePage,Count,Alias,TemplateName,FullUrl,Intro,MetaString)

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
			TemplateName=aryUserInfo(8)
			FullUrl=aryUserInfo(9)
			Intro=aryUserInfo(10)
			MetaString=aryUserInfo(11)

			If IsNull(Email) Or IsEmpty(Email) Or Len(Email)=0 Then Email="a@b.com"
			If IsNull(HomePage) Then HomePage=""
			If IsNull(Alias) Then Alias=""
			If IsNull(TemplateName) Then TemplateName=""

			LoadInfoByArray=True

		End If

		Call Filter_Plugin_TUser_LoadInfoByArray(ID,Name,Level,Password,Email,HomePage,Count,Alias,TemplateName,FullUrl,Intro,MetaString)

	End Function


	Function Edit(currentUser)

		Call Filter_Plugin_TUser_Edit(ID,Name,Level,Password,Email,HomePage,Count,Alias,TemplateName,FullUrl,Intro,MetaString,currentUser)

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(Level,"int",0)

		If ((Level<1) Or (Level>5)) Then Call ShowError(16)
		If (Name="") Then Call ShowError(7)
		If Len(Name) >ZC_USERNAME_MAX Then Call ShowError(7)
		If Not CheckRegExp(Name,"[username]") Then Call ShowError(7)

		Email=FilterSQL(Email)
		HomePage=FilterSQL(HomePage)

		Email=TransferHTML(Email,"[html-format]")
		HomePage=TransferHTML(HomePage,"[html-format]")

		Alias=TransferHTML(Alias,"[directory&file]")
		If Left(Alias,1)="/" Then Alias=Right(Alias,Len(Alias)-1)
		If Right(Alias,1)="/" Then Alias=Left(Alias,Len(Alias)-1)
		Alias=FilterSQL(Alias)

		TemplateName=UCase(FilterSQL(TemplateName))
		If TemplateName="CATALOG" Then TemplateName=""

		If Len(Email)=0 Then Call ShowError(29)
		If Len(Email)>ZC_EMAIL_MAX Then Call ShowError(29)
		If Len(HomePage)>ZC_HOMEPAGE_MAX Then Call ShowError(29)

		If Not CheckRegExp(Email,"[email]") Then Call ShowError(29)
		IF Len(HomePage)>0 Then
			If Not CheckRegExp(HomePage,"[homepage]") Then Call ShowError(30)
		End If

		Intro=FilterSQL(Intro)

		If ID=0 Then

			If Level <= currentUser.Level Then ShowError(6)
			If Len(PassWord)<>32 Then Call ShowError(55)

			Dim objRS2
			Set objRS2 = objConn.execute ("SELECT [mem_id] FROM [blog_Member] WHERE [mem_Name]='" & Name & "' ")
			If (Not objRS2.bof) And (Not objRS2.eof) Then
				Call ShowError(62)
			End If
			Set objRS2=Nothing

			objConn.Execute("INSERT INTO [blog_Member]([mem_Level],[mem_Name],[mem_PassWord],[mem_Email],[mem_HomePage],[mem_Url],[mem_Guid],[mem_Intro],[mem_Template],[mem_Meta]) VALUES ("&Level&",'"&Name&"','"&PassWord&"','"&Email&"','"&HomePage&"','"&Alias&"','"&Guid&"','"&Intro&"','"&TemplateName&"','"&MetaString&"')")

			Dim objRS
			Set objRS=objConn.Execute("SELECT MAX([mem_ID]) FROM [blog_Member]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
			End If
			Set objRS=Nothing

		Else

			If (ID=currentUser.ID) And (Level <> currentUser.Level) Then ShowError(6)
			If (ID<>currentUser.ID) And (Level <= currentUser.Level) Then ShowError(6)

			If ID>0 Then
				Dim objRS3
				Set objRS3 = objConn.execute ("SELECT [mem_id] FROM [blog_Member] WHERE [mem_Name]='" & Name & "' AND [mem_ID]<>" & ID)
				If (Not objRS3.bof) And (Not objRS3.eof) Then
					Call ShowError(62)
				End If
				Set objRS3=Nothing
			End If


			Dim targetUser
			Set targetUser=New TUser
			If targetUser.LoadInfobyID(ID) Then

				If Len(PassWord)=0 Then
					PassWord=targetUser.PassWord
				End If

				If Len(PassWord)<>32 Then Call ShowError(55)

				objConn.Execute("UPDATE [blog_Member] SET [mem_Level]="&Level&",[mem_Name]='"&Name&"',[mem_PassWord]='"&PassWord&"',[mem_Email]='"&Email&"',[mem_HomePage]='"&HomePage&"',[mem_Url]='"&Alias&"',[mem_Intro]='"&Intro&"',[mem_Template]='"&TemplateName&"',[mem_Meta]='"&MetaString&"' WHERE [mem_ID]="&ID)

				If Name <> targetUser.Name Then
					objConn.Execute("UPDATE [blog_Comment] SET [comm_Author]='"&Name&"' WHERE [comm_AuthorID]="&ID)
				End If
				If Email <> targetUser.Email Then
					objConn.Execute("UPDATE [blog_Comment] SET [comm_Email]='"&Email&"' WHERE [comm_AuthorID]="&ID)
				End If

			End If

			Dim tmpClsID
			tmpClsID=MD5(BlogPath & ZC_BLOG_CLSID_ORIGINAL)
			Application.Lock
			Application(tmpClsID & "QUICKLOGIN_ID" & Session(tmpClsID & "quicklogin_id"))=Empty
			Application(tmpClsID & "QUICKLOGIN_ID" & ID)=Empty
			Application.UnLock
			Session(tmpClsID & "quicklogin")=Empty
			Session(tmpClsID & "quicklogin_id")=Empty


		End If

		Edit=True

	End Function


	Function Register()

		Dim currentUser
		Set currentUser=BlogUser

		Call Filter_Plugin_TUser_Register(ID,Name,Level,Password,Email,HomePage,Count,Alias,TemplateName,FullUrl,Intro,MetaString,currentUser)

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(Level,"int",0)

		PassWord=MD5(Password & Guid)

		If (Level<=1) Then Call ShowError(16)
		If (Name="") Then Call ShowError(7)
		If Len(Name) >ZC_USERNAME_MAX Then Call ShowError(7)
		If Not CheckRegExp(Name,"[username]") Then Call ShowError(7)

		Email=FilterSQL(Email)
		HomePage=FilterSQL(HomePage)

		Email=TransferHTML(Email,"[html-format]")
		HomePage=TransferHTML(HomePage,"[html-format]")

		Alias=TransferHTML(Alias,"[directory&file]")
		If Left(Alias,1)="/" Then Alias=Right(Alias,Len(Alias)-1)
		If Right(Alias,1)="/" Then Alias=Left(Alias,Len(Alias)-1)
		Alias=FilterSQL(Alias)

		TemplateName=UCase(FilterSQL(TemplateName))
		If TemplateName="CATALOG" Then TemplateName=""

		If Len(Email)=0 Then Call ShowError(29)
		If Len(Email)>ZC_EMAIL_MAX Then Call ShowError(29)
		If Len(HomePage)>ZC_HOMEPAGE_MAX Then Call ShowError(29)

		If Not CheckRegExp(Email,"[email]") Then Call ShowError(30)
		IF Len(HomePage)>0 Then
			If Not CheckRegExp(HomePage,"[homepage]") Then Call ShowError(30)
		End If


		If ID=0 Then

			If Level <= 1 Then ShowError(6)
			If Len(PassWord)<>32 Then Call ShowError(55)

			Dim objRS2
			Set objRS2 = objConn.execute ("SELECT [mem_id] FROM [blog_Member] WHERE [mem_Name]='" & Name & "' ")
			If (Not objRS2.bof) And (Not objRS2.eof) Then
				Call ShowError(62)
			End If
			Set objRS2=Nothing

			objConn.Execute("INSERT INTO [blog_Member]([mem_Level],[mem_Name],[mem_PassWord],[mem_Email],[mem_HomePage],[mem_Url],[mem_Guid],[mem_Intro],[mem_Meta]) VALUES ("&Level&",'"&Name&"','"&PassWord&"','"&Email&"','"&HomePage&"','"&Alias&"','"&Guid&"','"&Intro&"','"&MetaString&"')")

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

		Call Filter_Plugin_TUser_Del(ID,Name,Level,Password,Email,HomePage,Count,Alias,TemplateName,FullUrl,Intro,MetaString,currentUser)

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

		Set objRS=objConn.Execute("SELECT ul_id FROM [blog_UpLoad] WHERE [ul_AuthorID] =" & ID)
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

	Function GetIDbyName(n)

		n=FilterSQL(n)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [mem_ID] FROM [blog_Member] WHERE [mem_Name]='"& n & "'")
		If (Not objRS.Bof) And (Not objRS.Eof) Then
			GetIDbyName=CLng(objRS("mem_ID"))
		End If

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

		ReCount=0
		ExID=-1

		Set Meta=New TMeta

	End Sub


End Class
'*********************************************************




'*********************************************************
' 目的：    定义TComment类
' 输入：    无
' 返回：    无
'*********************************************************
Class TComment

	Public ID
	Public log_ID
	Public ParentID

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
	Public IsCheck
	Public Meta

	Public IsThrow '此值为True时,系统不会保存的,会直接扔出去.

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property


	Private FAvatar
	Public Property Get Avatar
		'plugin node
		bAction_Plugin_TComment_Avatar=False
		For Each sAction_Plugin_TComment_Avatar in Action_Plugin_TComment_Avatar
			If Not IsEmpty(sAction_Plugin_TComment_Avatar) Then Call Execute(sAction_Plugin_TComment_Avatar)
			If bAction_Plugin_TComment_Avatar=True Then Exit Property
		Next

		If FAvatar="" Then FAvatar=Users(AuthorID).Avatar

		Avatar=FAvatar

	End Property


	Public html

	Public Property Get HomePageForAntiSpam
		HomePageForAntiSpam=URLEncodeForAntiSpam(HomePage)
	End Property


	Public Property Get SafeEmail
		Dim s
		If Email="" Then s="null@null.com"
		SafeEmail=Replace(s,"@","[AT]")
	End Property


	Public Property Get EmailMD5
		If AuthorID>0 Then
			EmailMD5=Users(AuthorID).EmailMD5
		Else
			If Email="" Or IsNull(Email)=True Then
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
			FirstContact="mailto:" & SafeEmail
		End If
	End Property

	Public Property Get HtmlContent
		'HtmlContent=TransferHTML(UBBCode(Content,"[font][face]"),"[enter][nofollow]")
		HtmlContent=TransferHTML(UBBCode(Content & Reply,"[link][link-antispam][font][face][typeset]"),"[enter][nofollow]")
	End Property


	Public Function Post()

		Call Filter_Plugin_TComment_Post(ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,IsCheck,MetaString)

		If IsThrow=True Then Post=True:Exit Function

		If IP="" Then
			IP=GetReallyIP()
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
		Call CheckParameter(IsCheck,"bool",False)

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
			objConn.Execute("INSERT INTO [blog_Comment]([log_ID],[comm_AuthorID],[comm_Author],[comm_Content],[comm_Email],[comm_HomePage],[comm_IP],[comm_PostTime],[comm_Agent],[comm_Reply],[comm_LastReplyIP],[comm_LastReplyTime],[comm_ParentID],[comm_IsCheck],[comm_Meta]) VALUES ("&log_ID&","&AuthorID&",'"&Author&"','"&Content&"','"&Email&"','"&HomePage&"','"&IP&"','"&PostTime&"','"&Agent&"','"&Reply&"','"&LastReplyIP&"','"&LastReplyTime&"','"&ParentID&"',"&CLng(IsCheck)&",'"&MetaString&"')")
			Set objRS=objConn.Execute("SELECT MAX([comm_ID]) FROM [blog_Comment]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
			End If
			Set objRS=Nothing
		Else
			objConn.Execute("UPDATE [blog_Comment] SET [log_ID]="&log_ID&", [comm_AuthorID]="&AuthorID&",[comm_Author]='"&Author&"',[comm_Content]='"&Content&"',[comm_Email]='"&Email&"',[comm_HomePage]='"&HomePage&"',[comm_IP]='"&IP&"',[comm_PostTime]='"&PostTime&"',[comm_Agent]='"&Agent&"',[comm_Reply]='"&Reply&"',[comm_LastReplyIP]='"&LastReplyIP&"',[comm_LastReplyTime]='"&LastReplyTime&"',[comm_ParentID]='"&ParentID&"',[comm_IsCheck]="&CLng(IsCheck)&",[comm_Meta]='"&MetaString&"' WHERE [comm_ID] =" & ID)
		End If

		Post=True

	End Function


	Public Function Del()

		Call Filter_Plugin_TComment_Del(ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,IsCheck,MetaString)

		Call CheckParameter(ID,"int",0)
		If (ID=0) Then Del=False:Exit Function
		objConn.Execute("DELETE FROM [blog_Comment] WHERE [comm_ID] =" & ID)
		Del=True
	End Function


	Public Function LoadInfoByID(comm_ID)

		Call CheckParameter(comm_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [comm_ID],[log_ID],[comm_AuthorID],[comm_Author],[comm_Content],[comm_Email],[comm_HomePage],[comm_PostTime],[comm_IP],[comm_Agent],[comm_Reply],[comm_LastReplyIP],[comm_LastReplyTime],[comm_ParentID],[comm_IsCheck],[comm_Meta] FROM [blog_Comment] WHERE [comm_ID]=" & comm_ID)

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
			IsCheck=objRS("comm_IsCheck")
			MetaString=objRS("comm_Meta")

			LoadInfoByID=True

		End If

		objRS.Close
		Set objRS=Nothing

		If IsNull(HomePage) Then HomePage=""

		Call Filter_Plugin_TComment_LoadInfoByID(ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,IsCheck,MetaString)

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
			IsCheck=aryCommInfo(14)
			MetaString=aryCommInfo(15)

		End If

		If IsNull(HomePage) Then HomePage=""

		LoadInfoByArray=True

		Call Filter_Plugin_TComment_LoadInfoByArray(ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,IsCheck,MetaString)

	End Function


	Public Function MakeTemplate(strC)

		Dim html,i,j
		html=strC

		'plugin node
		Call Filter_Plugin_TComment_MakeTemplate_Template(html)

		Dim aryTemplateTagsName()
		Dim aryTemplateTagsValue()

		ReDim aryTemplateTagsName(27)
		ReDim aryTemplateTagsValue(27)

		If ParentID="" Then ParentID=0

		Dim s
		If AuthorID>0 Then
			Call GetUsersbyUserIDList(AuthorID)
			If Users(AuthorID).Alias="" Then
				s=Users(AuthorID).Name
			Else
				s=Users(AuthorID).Alias
			End If
		Else
			s=Author
		End If

		aryTemplateTagsName(  1)="article/comment/id"
		aryTemplateTagsValue( 1)=ID
		aryTemplateTagsName(  2)="article/comment/name"
		aryTemplateTagsValue( 2)=s
		aryTemplateTagsName(  3)="article/comment/url"
		aryTemplateTagsValue( 3)=HomePage
		aryTemplateTagsName(  4)="article/comment/urlencoder"
		aryTemplateTagsValue( 4)=TransferHTML(HomePageForAntiSpam,"[anti-zc_blog_host]")
		aryTemplateTagsName(  5)="article/comment/email"
		aryTemplateTagsValue( 5)=SafeEmail
		aryTemplateTagsName(  6)="article/comment/posttime"
		aryTemplateTagsValue( 6)=PostTime
		aryTemplateTagsName(  7)="article/comment/content"
		aryTemplateTagsValue( 7)=HtmlContent & "<!--rev"&id&"-->" & "<a style=""display:none;"" id=""AjaxCommentEnd"&id&"""></a>"
		aryTemplateTagsName(  8)="article/comment/count"
		aryTemplateTagsValue( 8)="<!--(count-->"& Count &"<!--count)-->"
		aryTemplateTagsName(  9)="article/comment/authorid"
		aryTemplateTagsValue( 9)=AuthorID
		aryTemplateTagsName( 10)="article/comment/firstcontact"
		aryTemplateTagsValue(10)=TransferHTML(FirstContact,"[anti-zc_blog_host]")
		aryTemplateTagsName( 11)="article/comment/emailmd5"
		aryTemplateTagsValue(11)=EmailMD5
		aryTemplateTagsName( 12)="article/comment/parentid"
		aryTemplateTagsValue(12)=ParentID
		aryTemplateTagsName( 13)="article/comment/avatar"
		aryTemplateTagsValue(13)=Avatar
		aryTemplateTagsName( 14)="article/comment/posttime/longdate"
		aryTemplateTagsValue(14)=FormatDateTime(PostTime,vbLongDate)
		aryTemplateTagsName( 15)="article/comment/posttime/shortdate"
		aryTemplateTagsValue(15)=FormatDateTime(PostTime,vbShortDate)
		aryTemplateTagsName( 16)="article/comment/posttime/longtime"
		aryTemplateTagsValue(16)=FormatDateTime(PostTime,vbLongTime)
		aryTemplateTagsName( 17)="article/comment/posttime/shorttime"
		aryTemplateTagsValue(17)=FormatDateTime(PostTime,vbShortTime)
		aryTemplateTagsName( 18)="article/comment/posttime/year"
		aryTemplateTagsValue(18)=Year(PostTime)
		aryTemplateTagsName( 19)="article/comment/posttime/month"
		aryTemplateTagsValue(19)=Right("0"&Month(PostTime),2)
		aryTemplateTagsName( 20)="article/comment/posttime/monthname"
		aryTemplateTagsValue(20)=ZVA_Month(Month(PostTime))
		aryTemplateTagsName( 21)="article/comment/posttime/day"
		aryTemplateTagsValue(21)=Right("0"&Day(PostTime),2)
		aryTemplateTagsName( 22)="article/comment/posttime/weekday"
		aryTemplateTagsValue(22)=Weekday(PostTime)
		aryTemplateTagsName( 23)="article/comment/posttime/weekdayname"
		aryTemplateTagsValue(23)=ZVA_Week(Weekday(PostTime))
		aryTemplateTagsName( 24)="article/comment/posttime/hour"
		aryTemplateTagsValue(24)=Right("0"&Hour(PostTime),2)
		aryTemplateTagsName( 25)="article/comment/posttime/minute"
		aryTemplateTagsValue(25)=Right("0"&Minute(PostTime),2)
		aryTemplateTagsName( 26)="article/comment/posttime/second"
		aryTemplateTagsValue(26)=Right("0"&Second(PostTime),2)
		aryTemplateTagsName( 27)="article/comment/agent"
		aryTemplateTagsValue(27)=Agent

		'plugin node
		Call Filter_Plugin_TComment_MakeTemplate_TemplateTags(aryTemplateTagsName,aryTemplateTagsValue)

		j=UBound(aryTemplateTagsName)
		For i=1 to j
			If IsNull(aryTemplateTagsValue(i))=True Then aryTemplateTagsValue(i)=""
			html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
		Next

		Call Filter_Plugin_TComment_MakeTemplate_Template_Succeed(html)

		MakeTemplate=html

	End Function

	Private Sub Class_Initialize()

		ID=0
		log_ID=0
		AuthorID=0
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
		IP=GetReallyIP()
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
		If Len(Excerpt)>ZC_ARTICLE_EXCERPT_MAX Then Excerpt=Left(Excerpt,ZC_ARTICLE_EXCERPT_MAX)&"..."
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
		aryTemplateTagsValue( 3)=TransferHTML(UrlForAntiSpam,"[anti-zc_blog_host]")
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
	Public Quote

	Public Meta

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property

	Public html

	Public AutoName
	Public FullPath
	Public IsManual

	Public Function LoadInfoByID(ul_ID)

		Call CheckParameter(ul_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [ul_ID],[ul_AuthorID],[ul_FileSize],[ul_FileName],[ul_PostTime],[ul_FileIntro],[ul_DirByTime],[ul_Quote],[ul_Meta] FROM [blog_UpLoad] WHERE [ul_ID]=" & ul_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("ul_ID")
			AuthorID=objRS("ul_AuthorID")
			FileSize=objRS("ul_FileSize")
			FileName=objRS("ul_FileName")
			PostTime=objRS("ul_PostTime")
			FileIntro=objRS("ul_FileIntro")
			DirByTime=objRS("ul_DirByTime")
			Quote=objRS("ul_Quote")
			MetaString=objRS("ul_Meta")

			'If IsNull(DirByTime) Or DirByTime="" Then DirByTime=False

			LoadInfobyID=True

		End If

		objRS.Close
		Set objRS=Nothing

		Call Filter_Plugin_TUpLoadFile_LoadInfoByID(ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Quote,Meta)

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
			Quote=aryULInfo(7)
			MetaString=aryULInfo(8)

		End If

		LoadInfoByArray=True

		Call Filter_Plugin_TUpLoadFile_LoadInfoByArray(ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Quote,Meta)

	End Function


	Public Function UpLoad_Form()

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

		UpLoad_Form=True

	End Function


	Public Function UpLoad()

		Call Filter_Plugin_TUpLoadFile_UpLoad(ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Quote,Meta)

		DirByTime=True

		If IsManual=False Then
			Call UpLoad_Form()
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

		If AutoName=True Then
			Randomize
			FileName=Year(GetTime(Now())) & Right("0"&Month(GetTime(Now())),2) & Right("0"&Day(GetTime(Now())),2) & Right("0"&Hour(GetTime(Now())),2) & Right("0"&Minute(GetTime(Now())),2) & Right("0"&Second(GetTime(Now())),2) & Int(9 * Rnd) & Int(9 * Rnd) & Int(9 * Rnd) & Int(9 * Rnd) & Right(FileName,Len(FileName)-InStrRev(FileName,".")+1)
		End If

		FileIntro=FilterSQL(FileIntro)

		Dim objRS
		Set objRS=objConn.Execute("SELECT * FROM [blog_UpLoad] WHERE [ul_FileName] = '" & FileName & "'")

		If Len(FileName)>255 Then FileName=Right(FileName,255)
		PostTime=GetTime(Now())

		objConn.Execute("INSERT INTO [blog_UpLoad]([ul_AuthorID],[ul_FileSize],[ul_FileName],[ul_PostTime],[ul_FileIntro],[ul_DirByTime],[ul_Quote],[ul_Meta]) VALUES ("& AuthorID &","& FileSize &",'"& FileName &"','"& PostTime &"','"&FileIntro&"',"&CLng(DirByTime)&",'"&Quote&"','"&MetaString&"')")

		Dim strUPLOADDIR

		strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now()))

		Call CreatDirectoryByCustomDirectory(strUPLOADDIR)

		FullPath=BlogPath & strUPLOADDIR &"/" & FileName


		If IsManual=False Then
			Call SaveFile()
		End If


		UpLoad=True

	End Function


	Public Function SaveFile()

		Dim objStreamFile
		Set objStreamFile = Server.CreateObject("ADODB.Stream")

		objStreamFile.Type = adTypeBinary
		objStreamFile.Mode = adModeReadWrite
		objStreamFile.Open
		objStreamFile.Write Stream

		objStreamFile.SaveToFile FullPath,adSaveCreateOverWrite
		objStreamFile.Close

		SaveFile=True

	End Function



	Public Function Del()

		Call Filter_Plugin_TUpLoadFile_Del(ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Quote,Meta)

		Call CheckParameter(ID,"int",0)

		Dim objRS,strFilePath

		Set objRS=objConn.Execute("SELECT * FROM [blog_UpLoad] WHERE [ul_ID] = " & ID)

		If (Not objRS.bof) And (Not objRS.eof) Then


			strFilePath = BlogPath & ZC_UPLOAD_DIRECTORY &"/" & objRS("ul_FileName")
			Call DelToFile(strFilePath)


			strFilePath = BlogPath & ZC_UPLOAD_DIRECTORY & "/" & Year(objRS("ul_PostTime")) & "/" & Month(objRS("ul_PostTime")) &"/" & objRS("ul_FileName")
			Call DelToFile(strFilePath)

			objConn.Execute("DELETE FROM [blog_UpLoad] WHERE [ul_ID] =" & ID)

		Else

			Exit Function

		End If

		objRS.Close
		Set objRS=Nothing

		Del=True

	End Function

	Public Property Get FullUrl

		Dim strUPLOADDIR

		strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now()))

		FullUrl=BlogHost & strUPLOADDIR & "/" & FileName

	End Property

	Private Sub Class_Initialize()

		ID=0
		AuthorID=0
		Set Meta=New TMeta
		AutoName=False
		IsManual=False


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
	Public FullUrl
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
	Public Property Get EncodeIntro
		EncodeIntro = Server.URLEncode(Intro)
	End Property

	Private Ftemplate
	Public Property Let Template(strFileName)
		Ftemplate=GetTemplate("TEMPLATE_" & strFileName)
	End Property
	Public Property Get Template
		If Ftemplate<>"" Then
			Template = Ftemplate
			Exit Property
		Else
			If TemplateName<>"" Then
				Dim s
				s=GetTemplate("TEMPLATE_" &TemplateName)
				If s<>"" Then
					Ftemplate = s
				Else
					Ftemplate=GetTemplate("TEMPLATE_CATALOG")
				End If
			Else
				Ftemplate=GetTemplate("TEMPLATE_CATALOG")
			End If
			Template = Ftemplate
		End If
	End Property


	Public Function GetDefaultTemplateName
		If TemplateName<>"" Then
			GetDefaultTemplateName=TemplateName
		Else
			GetDefaultTemplateName="CATALOG"
		End If
	End Function


	Public Property Get FullPath
		FullPath=ParseCustomDirectoryForPath(FullRegex,ZC_STATIC_DIRECTORY,"","","","","",ID,Name,StaticName)
	End Property

	Public Property Get Url

		'plugin node
		bAction_Plugin_TTag_Url=False
		For Each sAction_Plugin_TTag_Url in Action_Plugin_TTag_Url
			If Not IsEmpty(sAction_Plugin_TTag_Url) Then Call Execute(sAction_Plugin_TTag_Url)
			If bAction_Plugin_TTag_Url=True Then Exit Property
		Next

		Url =ParseCustomDirectoryForUrl(FullRegex,ZC_STATIC_DIRECTORY,"","","","","",ID,EncodeName,StaticEncodeName)
		If Right(Url,12)="default.html" Then Url=Left(Url,Len(Url)-12)

		Url=Replace(Replace(Url,"//","/"),":/","://",1,1)

		Call Filter_Plugin_TTag_Url(Url)

	End Property


	Public Property Get StaticName
		If IsNull(Intro) Or IsEmpty(Intro) Or Intro="" Then
			StaticName = Name
		Else
			StaticName = Intro
		End If
	End Property

	Public Property Get StaticEncodeName
		If IsNull(Intro) Or IsEmpty(Intro) Or Intro="" Then
			StaticEncodeName = EncodeName
		Else
			StaticEncodeName = EncodeIntro
		End If
	End Property


	Private Ffullregex
	Public Property Let FullRegex(s)
		Ffullregex=s
	End Property
	Public Property Get FullRegex
		If Ffullregex<>"" Then
			FullRegex=Ffullregex
		Else
			FullRegex=ZC_TAGS_REGEX
		End If
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
		RssUrl = BlogHost & "feed.asp?tags=" & ID
	End Property

	Public ReCount

	Public Function Post()

		Call Filter_Plugin_TTag_Post(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString)

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(Order,"int",0)
		Call CheckParameter(ParentID,"int",0)

		Name=FilterSQL(Name)
		Name=TransferHTML(Name,"[normalname]")
		If Len(Name)=0 Then Post=False:Exit Function

		Intro=FilterSQL(Intro)
		Alias=FilterSQL(Alias)
		'Intro=TransferHTML(Intro,"[html-format]")
		If Intro="" Then
			Intro=Replace(Name," ","_")
		End If

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
		Set objRS=objConn.Execute("SELECT [tag_ID],[tag_Name],[tag_Intro],[tag_Order],[tag_Count],[tag_ParentID],[tag_URL],[tag_Template],[tag_FullUrl],[tag_Meta] FROM [blog_Tag] WHERE [tag_ID]=" & tag_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("tag_ID")
			Name=objRS("tag_Name")
			Intro=objRS("tag_Intro")
			Order=objRS("tag_Order")
			Count=objRS("tag_Count")
			ParentID=objRS("tag_ParentID")
			Alias=objRS("tag_URL")
			TemplateName=objRS("tag_Template")
			FullUrl=objRS("tag_FullUrl")
			MetaString=objRS("tag_Meta")

			LoadInfoByID=True

		End If

		objRS.Close
		Set objRS=Nothing

		If IsNull(Intro) Then Intro=""

		Call Filter_Plugin_TTag_LoadInfoByID(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString)

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
			FullUrl=aryTagInfo(8)
			MetaString=aryTagInfo(9)
		End If

		If IsNull(Intro) Then Intro=""
		If IsNull(Alias) Then Alias=""

		LoadInfoByArray=True

		Call Filter_Plugin_TTag_LoadInfoByArray(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString)

	End Function


	Public Function Del()

		Call Filter_Plugin_TTag_Del(ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString)

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

		ID=0
		Order=0
		ParentID=0
		Count=0
		ReCount=0

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
		'If(Len(wfw_comment)>0) Then
		'	objXMLitem.AppendChild(objXMLdoc.createElement("wfw:comment"))
		'	objXMLitem.selectSingleNode("wfw:comment").text=wfw_comment
		'End If
		If(Len(wfw_commentRss)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("wfw:commentRss"))
			objXMLitem.selectSingleNode("wfw:commentRss").text=wfw_commentRss
		End If
		'If(Len(trackback_ping)>0) Then
		'	objXMLitem.AppendChild(objXMLdoc.createElement("trackback:ping"))
		'	objXMLitem.selectSingleNode("trackback:ping").text=trackback_ping
		'End If

		objXMLchannel.AppendChild(objXMLitem)

		AddItem=True

	End Function

	Public Function Execute()

		'Response.ContentType = "text/html"
		Response.ContentType = "text/xml"
		Response.Clear
		Call Filter_Plugin_TNewRss2Export_PreExeOrSave(Me,"Execute")
		Response.Write xml

		Execute=True

	End Function

	Public Function SaveToFile(strFileName)
		On Error Resume Next
		Call Filter_Plugin_TNewRss2Export_PreExeOrSave(Me,"SaveToFile")
		objXMLdoc.save(strFileName)

		SaveToFile=True
		Err.Clear
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

		ParseDateForRFC822 = dtmWeekDay & ", " & dtmDay &" " & dtmMonth & " " & dtmYear & " " & dtmHours & ":" & dtmMinutes & ":" & dtmSeconds & IIF(CLng(TimeZone)=0,""," " & TimeZone)

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

		'Set objPI = objXMLdoc.createProcessingInstruction("xml-stylesheet","type=""text/xsl"" href=""zb_system/css/rss.xslt""")
		'objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
		'Set objPI = Nothing

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
Dim meta_split_string_1
Dim meta_split_string_2
meta_split_string_1=Chr(1)
meta_split_string_2=Chr(2)
Class TMeta

	Public Names()
	Public Values()

	Public Property Get Count
		Count = UBound(names)
	End Property

	Private Function Save()

		Dim n,v
		If UBound(names)>0 Then

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

	Private Function Load(s)

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


	Private Function LoadByArray(n,v)

		Dim i
		ReDim Names(UBound(n))
		ReDim Values(UBound(n))
		For i=0 To UBound(n)
			Names(i)=n(i)
			Values(i)=v(i)
		Next

	End Function


	Public Function SetValue(name,value)

		name=Trim(name)

		If IsEmpty(name) Or IsNull(name) Then Exit Function

		name=Replace(name,meta_split_string_1,"")
		name=Replace(name,meta_split_string_2,"")

		If IsNull(value)=True Then value=""

		Dim n,i
		i=0
		For Each n In names
			If LCase(n)=LCase(name) Then
				values(i)=vbsescape(value)
				'If values(i)="undefined" Then values(i)=""
				Exit function
			End If
			i=i+1
		Next

		i=UBound(names)

		ReDim Preserve Names(i+1)
		ReDim Preserve Values(i+1)

		Names(i+1)=name
		Values(i+1)=vbsescape(value)
		'If values(i+1)="undefined" Then values(i+1)=""

	End Function

	Public Function GetValue(name)

		Dim n,i
		i=0
		For Each n In names

			If LCase(n)=LCase(name) Then
				GetValue = vbsunescape(values(i))
				Exit function
			End If
			i=i+1
		Next

		GetValue = Empty
	End Function

	Public Function Remove(name)

		Dim n()
		Dim v()
		Dim b

		ReDim n(UBound(names))
		ReDim v(UBound(names))

		Dim i,j
		j=0
		For i=0 To UBound(names)
			If LCase(names(i))=LCase(name) Then
				b=True
			Else
				n(j)=names(i)
				v(j)=values(i)
				j=j+1
			End If
		Next

		If b=True Then

			ReDim names(j-1)
			ReDim values(j-1)

			For i=0 To j-1
				names(i)=n(i)
				values(i)=v(i)
			Next

		End If

	End Function

	Public Function Exists(name)

		If Name="" Then Exit Function

		Dim s
		For Each s In names
			If LCase(s)=LCase(name) Then
				Exists=True
				Exit Function
			End If
		Next

		Exists=False

	End Function

	Private Sub Class_Initialize()

		ReDim Names(0)
		ReDim Values(0)

	End Sub

End Class
'*********************************************************










'*********************************************************
' 目的：    定义TConfig
' 输入：    无
' 返回：    无
'*********************************************************
Class TConfig

	Private Name

	Public Meta

	Public Property Get Count
		Count = Meta.Count
	End Property

	Public Function Save()

		Dim n,s
		n=FilterSQL(Name)
		s=FilterSQL(Meta.SaveString)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [conf_Name] FROM [blog_Config] WHERE [conf_Name]='"&n&"'" )
		If (Not objRS.bof) And (Not objRS.eof) Then
			objConn.Execute("UPDATE [blog_Config] SET [conf_Value]='"&s&"' WHERE [conf_Name]='"&n&"'")
		Else
			objConn.Execute("INSERT INTO [blog_Config]([conf_Name],[conf_Value]) VALUES ('"&n&"','"&s&"')")
		End If
		Set objRS=Nothing

		Call ConfigMetas.SetValue(n,s)

	End Function

	Public Function Delete

		objConn.Execute("DELETE FROM [blog_Config] WHERE [conf_Name]='"&Name&"'")

		Call ConfigMetas.Remove(Name)

	End Function

	Public Function Load(configname)

		Call GetConfigs()

		Name=configname

		Dim s
		For Each s In ConfigMetas.Names
			If LCase(s)=LCase(Name) Then
				Meta.LoadString=ConfigMetas.GetValue(Name)
			End If
		Next

	End Function


	Public Function Write(name,value)
		If TypeName(value)="Boolean" And UCase(value)=UCase("true") Then value="True"
		If TypeName(value)="Boolean" And  UCase(value)=UCase("false") Then value="False"
		Write=Meta.SetValue(name,value)

	End Function

	Public Function Read(name)

		Read=Meta.GetValue(name)

	End Function

	Public Function Remove(name)

		Remove=Meta.Remove(name)

	End Function


	Public Function Exists(name)

		Exists=Meta.Exists(name)

	End Function


	Private Sub Class_Initialize()

		Name=Empty

		Set Meta=New TMeta

	End Sub

End Class
'*********************************************************







'*********************************************************
' 目的：    定义模块功能类
' 输入：    无
' 返回：    无
'*********************************************************
Class TFunction

	Public ID
	Public Name
	Public FileName
	Public Order
	Public Content
	Public IsHidden
	Public SidebarID
	Public HtmlID
	Public Ftype 'div or ul
	Public MaxLi
	Public Meta
	Public Source
	Public ViewType
	Public InDataBase
	Public IsHideTitle

	Public Property Get IsDisplay
		IsDisplay=Not IsHidden
	End Property

	Public Property Get IsSystem
		If Source="system" Then
			IsSystem=True
		Else
			IsSystem=False
		End If
	End Property
	Public Property Let IsSystem(s)
		If s=True Then
			Source="system"
		End If
	End Property

	Public Property Get IsUsers
		If Source="users" Then
			IsUsers=True
		Else
			IsUsers=False
		End If
	End Property
	Public Property Let IsUsers(s)
		If s=True Then
			Source="users"
		End If
	End Property

	Public Property Get IsPlugin
		If InStr(Source,"plugin_")>0 Then
			IsPlugin=True
		Else
			IsPlugin=False
		End If
	End Property
	Public Property Let IsPlugin(s)
		Source="plugin_"&s
	End Property


	Public Property Get IsTheme
		If InStr(Source,"theme_")>0 Then
			IsTheme=True
		Else
			IsTheme=False
		End If
	End Property
	Public Property Let IsTheme(s)
		Source="theme_"&s
	End Property

	Public Property Get SourceType
		If Source="system" Then
			SourceType="system"
		End If
		If Source="users" Then
			SourceType="users"
		End If
		If InStr(Source,"plugin_")>0 Then
			SourceType="plugin"
		End If
		If InStr(Source,"theme_")>0 Then
			SourceType="theme"
		End If
		If Source="other" Then
			SourceType="other"
		End If
	End Property

	Public Property Get IsOther
		If Source="other" Then
			IsOther=True
		Else
			IsOther=False
		End If
	End Property
	Public Property Let IsOther(s)
		If s=True Then
			Source="other"
		End If
	End Property

	Public Property Get AppName
		If SourceType="plugin" Then
			AppName=Right(Source,Len(Source)-Len("plugin_"))
		ElseIf SourceType="theme" Then
			AppName=Right(Source,Len(Source)-Len("theme_"))
		Else
			AppName=""
		End If
	End Property

	Public Property Get MetaString
		MetaString=Meta.SaveString
	End Property
	Public Property Let MetaString(s)
		Meta.LoadString=s
	End Property

	Public Function Post()

		Call Filter_Plugin_TFunction_Post(ID,Name,FileName,Order,Content,IsHidden,SidebarID,HtmlID,Ftype,MaxLi,Source,ViewType,IsHideTitle,MetaString)

		Call CheckParameter(ID,"int",0)
		Call CheckParameter(Order,"int",0)
		Call CheckParameter(SidebarID,"int",1)
		Call CheckParameter(IsHidden,"bool",False)
		Call CheckParameter(IsHideTitle,"bool",False)
		Call CheckParameter(MaxLi,"int",0)

		Name=FilterSQL(Name)
		FileName=TransferHTML(LCase(FilterSQL(FileName)),"[delspace][filename][normalname]")
		FileName=Replace(FileName,".","")
		HtmlID=TransferHTML(FilterSQL(HtmlID),"[delspace][filename]")
		Source=FilterSQL(Source)
		ViewType=FilterSQL(ViewType)

		If Name="" Then
			Name="Function"
		End If

		If FileName="" Then
			If ID=0 Then
				FileName="function"& GetNewID
			Else
				FileName="function" & ID
			End If
		End If

		If HtmlID="" Then
			If ID=0 Then
				HtmlID="divFunction" & GetNewID
			Else
				HtmlID="divFunction" & ID
			End If
		End If

		If Order=0 Then
			Order=GetNewOrder
		End If

		If Source="" Then
			Source="users"
		End If

		Name=Left(Name,50)
		FileName=Left(FileName,50)
		HtmlID=Left(HtmlID,50)

		If Ftype<>"div" And Ftype<>"ul" Then Ftype="div"

		Dim sContent
		sContent=Content
		Content=FilterSQL(Content)
		Content=TransferHTML(Content,"[anti-zc_blog_host]")


		If ID=0 Or (InDataBase=False And Source<>"other") Then
			objConn.Execute("INSERT INTO [blog_Function]([fn_Name],[fn_FileName],[fn_Order],[fn_Content],[fn_IsHidden],[fn_SidebarID],[fn_HtmlID],[fn_Ftype],[fn_MaxLi],[fn_Source],[fn_ViewType],[fn_IsHideTitle],[fn_Meta]) VALUES ('"&Name&"','"&FileName&"',"&Order&",'"&Content&"',"&CLng(IsHidden)&","&SidebarID&",'"&HtmlID&"','"&Ftype&"',"&MaxLi&",'"&Source&"','"&ViewType&"',"&CLng(IsHideTitle)&",'"&MetaString&"')")

			Dim objRS
			Set objRS=objConn.Execute("SELECT MAX([fn_ID]) FROM [blog_Function]")
			If (Not objRS.bof) And (Not objRS.eof) Then
				ID=objRS(0)
				InDataBase=True
			End If

		Else
			If InDataBase Then
				objConn.Execute("UPDATE [blog_Function] SET [fn_Name]='"&Name&"',[fn_FileName]='"&FileName&"',[fn_Order]="&Order&",[fn_Content]='"&Content&"',[fn_IsHidden]="&CLng(IsHidden)&",[fn_SidebarID]="&SidebarID&",[fn_HtmlID]='"&HtmlID&"',[fn_Ftype]='"&Ftype&"',[fn_MaxLi]="&MaxLi&",[fn_Source]='"&Source&"',[fn_ViewType]='"&ViewType&"',[fn_IsHideTitle]="&CLng(IsHideTitle)&",[fn_Meta]='"&MetaString&"' WHERE [fn_ID] =" & ID)
			End If
		End If

		Content=sContent

		Post=True

	End Function


	Public Function LoadInfoByID(fn_ID)

		Call CheckParameter(fn_ID,"int",0)

		Dim objRS
		Set objRS=objConn.Execute("SELECT [fn_ID],[fn_Name],[fn_FileName],[fn_Order],[fn_Content],[fn_IsHidden],[fn_SidebarID],[fn_HtmlID],[fn_Ftype],[fn_MaxLi],[fn_Source],[fn_ViewType],[fn_IsHideTitle],[fn_Meta] FROM [blog_Function] WHERE [fn_ID]=" & fn_ID)

		If (Not objRS.bof) And (Not objRS.eof) Then

			ID=objRS("fn_ID")
			Name=objRS("fn_Name")
			FileName=objRS("fn_FileName")
			Order=objRS("fn_Order")
			Content=objRS("fn_Content")
			IsHidden=objRS("fn_IsHidden")
			SidebarID=objRS("fn_SidebarID")
			HtmlID=objRS("fn_HtmlID")
			Ftype=objRS("fn_Ftype")
			MaxLi=objRS("fn_MaxLi")
			Source=objRS("fn_Source")
			ViewType=objRS("fn_ViewType")
			IsHideTitle=objRS("fn_IsHideTitle")
			MetaString=objRS("fn_Meta")

			LoadInfoByID=True
			InDataBase=True

		End If

		objRS.Close
		Set objRS=Nothing

		Call Filter_Plugin_TFunction_LoadInfoByID(ID,Name,FileName,Order,Content,IsHidden,SidebarID,HtmlID,Ftype,MaxLi,Source,ViewType,IsHideTitle,MetaString)

	End Function



	Public Function LoadInfoByArray(aryCateInfo)

		If IsArray(aryCateInfo)=True Then
			ID=aryCateInfo(0)
			Name=aryCateInfo(1)
			FileName=aryCateInfo(2)
			Order=aryCateInfo(3)
			Content=aryCateInfo(4)
			IsHidden=aryCateInfo(5)
			SidebarID=aryCateInfo(6)
			HtmlID=aryCateInfo(7)
			Ftype=aryCateInfo(8)
			MaxLi=aryCateInfo(9)
			Source=aryCateInfo(10)
			ViewType=aryCateInfo(11)
			IsHideTitle=aryCateInfo(12)
			MetaString=aryCateInfo(13)
		End If

		LoadInfoByArray=True
		InDataBase=True

		Call Filter_Plugin_TFunction_LoadInfoByArray(ID,Name,FileName,Order,Content,IsHidden,SidebarID,HtmlID,Ftype,MaxLi,Source,ViewType,IsHideTitle,MetaString)

	End Function


	Public Function GetNewID()

		GetNewID=CLng(objConn.Execute("SELECT TOP 1 [fn_ID] FROM [blog_Function] ORDER BY [fn_ID] DESC")(0))+1

	End Function


	Public Function GetNewOrder()

		GetNewOrder=CLng(objConn.Execute("SELECT TOP 1 [fn_Order] FROM [blog_Function] ORDER BY [fn_Order] DESC")(0))+1

	End Function


	Public Function InSidebars(num)
		If num=1 Then InSidebars=InSidebar
		If num=2 Then InSidebars=InSidebar2
		If num=3 Then InSidebars=InSidebar3
		If num=4 Then InSidebars=InSidebar4
		If num=5 Then InSidebars=InSidebar5
	End Function


	Public Function InSidebar()
		If SidebarID=-1 Then InSidebar=False:Exit Function
		InSidebar=(Round(Right(SidebarID,1)/1)=1)
	End Function

	Public Function InSidebar2()
		If SidebarID=-1 Then InSidebar2=False:Exit Function
		InSidebar2=(Round(Right(SidebarID,2)/11)=1)
	End Function

	Public Function InSidebar3()
		If SidebarID=-1 Then InSidebar3=False:Exit Function
		InSidebar3=(Round(Right(SidebarID,3)/111)=1)
	End Function

	Public Function InSidebar4()
		If SidebarID=-1 Then InSidebar4=False:Exit Function
		InSidebar4=(Round(Right(SidebarID,4)/1111)=1)
	End Function

	Public Function InSidebar5()
		If SidebarID=-1 Then InSidebar5=False:Exit Function
		InSidebar5=(Round(Right(SidebarID,5)/11111)=1)
	End Function



	Public Function MakeTemplate(strFunction)

		Dim html,i,j,s,RE
		html=strFunction

		Set RE = New RegExp
		RE.IgnoreCase = True
		RE.Global = True

		If IsHideTitle=True Then
			RE.Pattern = "<#template:function_title:begin#>(.|\n)*<#template:function_title:end#>"
			html = RE.Replace(html, "")
		Else
			html=Replace(html,"<#template:function_title:begin#>","")
			html=Replace(html,"<#template:function_title:end#>","")
		End If

		If Ftype="div" Then
			s="<div><#CACHE_INCLUDE_" & UCase(FileName) & "#></div>"
			If ViewType="js" Then
				s="<div><#CACHE_INCLUDE_" & UCase(FileName) & "_JS#></div>"
			End If
			If ViewType="html" Then
				s="<div><#CACHE_INCLUDE_" & UCase(FileName) & "_HTML#></div>"
			End If
		End If

		If Ftype="ul" Then
			s="<ul><#CACHE_INCLUDE_" & UCase(FileName) & "#></ul>"
			If ViewType="js" Then
				s="<ul><#CACHE_INCLUDE_" & UCase(FileName) & "_JS#></ul>"
			End If
			If ViewType="html" Then
				s="<ul><#CACHE_INCLUDE_" & UCase(FileName) & "_HTML#></ul>"
			End If
		End If

		'plugin node


		Dim aryTemplateTagsName()
		Dim aryTemplateTagsValue()

		ReDim aryTemplateTagsName(5)
		ReDim aryTemplateTagsValue(5)

		aryTemplateTagsName(  1)="function/id"
		aryTemplateTagsValue( 1)=ID
		aryTemplateTagsName(  2)="function/name"
		aryTemplateTagsValue( 2)=Name
		aryTemplateTagsName(  3)="function/htmlid"
		aryTemplateTagsValue( 3)=HtmlID
		aryTemplateTagsName(  4)="function/content"
		aryTemplateTagsValue( 4)=s
		aryTemplateTagsName(  5)="function/filename"
		aryTemplateTagsValue( 5)=FileName


		'plugin node


		j=UBound(aryTemplateTagsName)
		For i=1 to j
			html=Replace(html,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
		Next

		MakeTemplate=html

	End Function


	Public Function Save()

		Call Post()
		Call SaveFile()

		Save=True

	End Function


	Public Function SaveFile()

		If Ftype="ul" And MaxLi>0 Then
			Dim i,b,j
			b=Split(Content,"</li>")
			If UBound(b)>0 then
				For i=0 To UBound(b)-1
					j=j+1
					b(i)=b(i) & "</li>"
					If j>Maxli Then
						b(i)=""
					End If
				Next
				Content=Join(b,"")
			End if
		End If

		If IsHidden=True THen
			Content=""
		Else
			Content=TransferHTML(Content,"[anti-zc_blog_host]")
		End If

		Call SaveToFile(BlogPath & "zb_users/include/"&FileName&".asp",Content,"utf-8",False)

		SaveFile=True

	End Function


	Public Function Del()

		Call Filter_Plugin_TFunction_Del(ID,Name,FileName,Order,Content,IsHidden,SidebarID,HtmlID,Ftype,MaxLi,Source,ViewType,IsHideTitle,MetaString)

		Call CheckParameter(ID,"int",0)

		If (ID=0) Then Del=False:Exit Function

		objConn.Execute("DELETE FROM [blog_Function] WHERE [fn_ID] =" & ID)

		InDataBase=False

		Call DelFile()

		Del=True

	End Function


	Public Function DelFile()

		Call DelToFile(BlogPath & "zb_users/include/" & FileName & ".asp")

	End Function



	Private Sub Class_Initialize()
		ID=0
		Ftype="div"
		SidebarID=1
		IsHidden=False
		IsHideTitle=False
		MaxLi=0
		Source="users"
		InDataBase=False
		Set Meta=New TMeta
	End Sub


End Class
'*********************************************************




'*********************************************************
'                        TCounter
'*********************************************************
Class TCounter

	Public ID
	Public IP
	Public Agent
	Public Content
	Public UserID
	Public PostTime
	Public PostData
	Public URL
	Public Referer
	Public AllRequestHeader
	Public Name

	Public Function LoadInfoById(nid)
		Call CheckParameter(nid,"int",0)
		Dim strSQL
		strSQL="SELECT [coun_ID],[coun_IP],[coun_Agent],[coun_Refer],[coun_PostTime],[coun_Content],[coun_UserID],[coun_PostData],[coun_URL],[coun_AllRequestHeader],[coun_logName] FROM [blog_Counter] WHERE [coun_ID]="&nid
		Dim objRs
		Set objRs=objConn.Execute(strsql)
		If (Not objRS.bof) And (Not objRS.eof) Then
			ID=objRs(0)
			IP=vbsunescape(objRs(1))
			Agent=vbsunescape(objRs(2))
			Referer=vbsunescape(objRs(3))
			PostTime=objRs(4)
			Content=vbsunescape(objRs(5))
			UserID=objRs(6)
			PostData=vbsunescape(objRs(7))
			URL=vbsunescape(objRs(8))
			AllRequestHeader=vbsunescape(objRs(9))
			Name=vbsunescape(objRs(10))
			LoadInfoById=True
		End If
	End Function

	Public Function LoadInfoByArray(objRs)

		If IsArray(objRs)=True Then

			ID=objRs(0)
			IP=vbsunescape(objRs(1))
			Agent=vbsunescape(objRs(2))
			Referer=vbsunescape(objRs(3))
			PostTime=objRs(4)
			Content=vbsunescape(objRs(5))
			UserID=objRs(6)
			PostData=vbsunescape(objRs(7))
			URL=vbsunescape(objRs(8))
			AllRequestHeader=vbsunescape(objRs(9))
			PostTime = Year(PostTime) & "-" & Month(PostTime) & "-" & Day(PostTime) & " " & Hour(PostTime) & ":" & Minute(PostTime) & ":" & Second(PostTime)
			Name=vbsunescape(objRs(10))
			LoadInfoByArray=True
		End If


	End Function

	Public Function Add(b,c)
		IP=GetReallyIP()
		Agent=Request.ServerVariables("HTTP_USER_AGENT")
		Referer=Request.ServerVariables("HTTP_REFERER")
		PostTime=FormatDateTime(Now)
		Content=c
		UserID=BlogUser.ID
		PostData=IIf(IIf(InStr(LCase(Request.ServerVariables("CONTENT_TYPE")),"multipart/form-data")>0,True,False),"Binary PostData",Request.Form)
		URL=GetUrl
		Name=b
		AllRequestHeader=Request.ServerVariables("ALL_HTTP")
		Dim j,k,i,m
		j=Array(IP,Agent,Referer,PostTime,Content,UserID,PostData,URL,AllRequestHeader,Name)
		k=Array("coun_IP","coun_Agent","coun_Refer","coun_PostTime","coun_Content","coun_UserID","coun_PostData","coun_URL","coun_AllRequestHeader","coun_logName")
		m="INSERT INTO [blog_Counter]("
		For i=0 To Ubound(k)-1
			m=m&"["&k(i)&"],"
		Next
		m=m&"["&k(i)&"]) VALUES("
		For i=0 To Ubound(j)-1
			m=m&"'"&IIf(i=3 or i=5,j(i),vbsescape(j(i)))&"',"
		Next
		m=m&"'"&vbsescape(j(i))&"'"
		m=m&")"
		objConn.Execute m

		ID=objConn.Execute("SELECT MAX([coun_ID]) FROM [blog_Counter]")(0)
	End Function

	Public Function GetUrl
		GetUrl=Request.ServerVariables("HTTP_METHOD")&": "
		GetUrl=GetUrl&IIf(LCase(Request.ServerVariables("HTTPS"))="off","http://","https://")
		GetUrl=GetUrl & Request.ServerVariables("SERVER_NAME")
		GetUrl=IIf(Request.ServerVariables("SERVER_PORT")<>80,GetUrl&":"&Request.ServerVariables("SERVER_PORT"),GetUrl)
		GetUrl=GetUrl&Request.ServerVariables("URL")
		GetUrl=IIf(Trim(Request.QueryString)<>"",GetUrl&"?"&Trim(Request.QueryString),GetUrl)
	End Function

	Function DelOld(interval, old)
		If ZC_MSSQL_ENABLE Then
			objConn.Execute("DELETE FROM [blog_Counter] WHERE DATEDIFF("&interval&",[coun_PostTime],getdate())>"&old)
		Else
			objConn.Execute("DELETE FROM [blog_Counter] WHERE DATEDIFF('"&interval&"',[coun_PostTime],now())>"&old)
		End If
	End Function
End Class
'*********************************************************


'----------------------------------------------------------
'**************  风声 ASP 无组件上传类 V2.11  *************
'作者：风声
'网站：http://www.fonshen.com
'邮件：webmaster@fonshen.com
'版权：版权全体,源代码公开,各种用途均可免费使用
'其他：有稍作改动
'**********************************************************
'----------------------------------------------------------
Class UpLoadClass

	Private m_TotalSize,m_MaxSize,m_FileType,m_SavePath,m_AutoSave,m_Error,m_Charset
	Private m_dicForm,m_binForm,m_binItem,m_strDate,m_lngTime
	Public	FormItem,FileItem

	Public Property Get Version
		Version="Fonshen ASP UpLoadClass Version 2.11"
	End Property

	Public Property Get Error
		Error=m_Error
	End Property

	Public Property Get Charset
		Charset=m_Charset
	End Property
	Public Property Let Charset(strCharset)
		m_Charset=strCharset
	End Property

	Public Property Get TotalSize
		TotalSize=m_TotalSize
	End Property
	Public Property Let TotalSize(lngSize)
		if isNumeric(lngSize) then m_TotalSize=Clng(lngSize)
	End Property

	Public Property Get MaxSize
		MaxSize=m_MaxSize
	End Property
	Public Property Let MaxSize(lngSize)
		if isNumeric(lngSize) then m_MaxSize=Clng(lngSize)
	End Property

	Public Property Get FileType
		FileType=m_FileType
	End Property
	Public Property Let FileType(strType)
		m_FileType=strType
	End Property

	Public Property Get SavePath
		SavePath=m_SavePath
	End Property
	Public Property Let SavePath(strPath)
		m_SavePath=Replace(strPath,chr(0),"")
	End Property

	Public Property Get AutoSave
		AutoSave=m_AutoSave
	End Property
	Public Property Let AutoSave(byVal Flag)
		select case Flag
			case 0,1,2: m_AutoSave=Flag
		end select
	End Property

	Private Sub Class_Initialize
		m_Error	   = -1
		m_Charset  = "gb2312"
		m_TotalSize= 0
		m_MaxSize  = 153600
		m_FileType = "jpg/gif"
		m_SavePath = ""
		m_AutoSave = 0
		Dim dtmNow : dtmNow = Date()
		m_strDate  = Year(dtmNow)&Right("0"&Month(dtmNow),2)&Right("0"&Day(dtmNow),2)
		m_lngTime  = Clng(Timer()*1000)
		Set m_binForm = Server.CreateObject("ADODB.Stream")
		Set m_binItem = Server.CreateObject("ADODB.Stream")
		Set m_dicForm = Server.CreateObject("Scripting.Dictionary")
		m_dicForm.CompareMode = 1
	End Sub

	Private Sub Class_Terminate
		m_dicForm.RemoveAll
		Set m_dicForm = nothing
		Set m_binItem = nothing
		m_binForm.Close()
		Set m_binForm = nothing
	End Sub

	Public Function Open()
		Open = 0
		if m_Error=-1 then
			m_Error=0
		else
			Exit Function
		end if
		Dim lngRequestSize : lngRequestSize=Request.TotalBytes
		if m_TotalSize>0 and lngRequestSize>m_TotalSize then
			m_Error=5
			Exit Function
		elseif lngRequestSize<1 then
			m_Error=4
			Exit Function
		end if

		Dim lngChunkByte : lngChunkByte = 102400
		Dim lngReadSize : lngReadSize = 0
		m_binForm.Type = 1
		m_binForm.Open()
		do
			m_binForm.Write Request.BinaryRead(lngChunkByte)
			lngReadSize=lngReadSize+lngChunkByte
			if  lngReadSize >= lngRequestSize then exit do
		loop
		m_binForm.Position=0
		Dim binRequestData : binRequestData=m_binForm.Read()

		Dim bCrLf,strSeparator,intSeparator
		bCrLf=ChrB(13)&ChrB(10)
		intSeparator=InstrB(1,binRequestData,bCrLf)-1
		strSeparator=LeftB(binRequestData,intSeparator)

		Dim strItem,strInam,strFtyp,strPuri,strFnam,strFext,lngFsiz
		Const strSplit="'"">"
		Dim strFormItem,strFileItem,intTemp,strTemp
		Dim p_start : p_start=intSeparator+2
		Dim p_end
		Do
			p_end = InStrB(p_start,binRequestData,bCrLf&bCrLf)-1
			m_binItem.Type=1
			m_binItem.Open()
			m_binForm.Position=p_start
			m_binForm.CopyTo m_binItem,p_end-p_start
			m_binItem.Position=0
			m_binItem.Type=2
			m_binItem.Charset=m_Charset
			strItem = m_binItem.ReadText()
			m_binItem.Close()
			intTemp=Instr(39,strItem,"""")
			strInam=Mid(strItem,39,intTemp-39)

			p_start = p_end + 4
			p_end = InStrB(p_start,binRequestData,strSeparator)-1
			m_binItem.Type=1
			m_binItem.Open()
			m_binForm.Position=p_start
			lngFsiz=p_end-p_start-2
			m_binForm.CopyTo m_binItem,lngFsiz


			if Instr(intTemp,strItem,"filename=""")<>0 then
			if not m_dicForm.Exists(strInam&"_From") then
				strFileItem=strFileItem&strSplit&strInam
				if m_binItem.Size<>0 then
					intTemp=intTemp+13
					strFtyp=Mid(strItem,Instr(intTemp,strItem,"Content-Type: ")+14)
					strPuri=Mid(strItem,intTemp,Instr(intTemp,strItem,"""")-intTemp)
					intTemp=InstrRev(strPuri,"\")
					strFnam=Mid(strPuri,intTemp+1)
					m_dicForm.Add strInam&"_Type",strFtyp
					m_dicForm.Add strInam&"_Name",strFnam
					m_dicForm.Add strInam&"_Path",Left(strPuri,intTemp)
					m_dicForm.Add strInam&"_Size",lngFsiz
					if Instr(strFnam,".")<>0 then
						strFext=Mid(strFnam,InstrRev(strFnam,".")+1)
					else
						strFext=""
					end if

					select case strFtyp
					case "image/jpeg","image/pjpeg","image/jpg"
						if Lcase(strFext)<>"jpg" then strFext="jpg"
						m_binItem.Position=3
						do while not m_binItem.EOS
							do
								intTemp = Ascb(m_binItem.Read(1))
							loop while intTemp = 255 and not m_binItem.EOS
							if intTemp < 192 or intTemp > 195 then
								m_binItem.read(Bin2Val(m_binItem.Read(2))-2)
							else
								Exit do
							end if
							do
								intTemp = Ascb(m_binItem.Read(1))
							loop while intTemp < 255 and not m_binItem.EOS
						loop
						m_binItem.Read(3)
						m_dicForm.Add strInam&"_Height",Bin2Val(m_binItem.Read(2))
						m_dicForm.Add strInam&"_Width",Bin2Val(m_binItem.Read(2))
					case "image/gif"
						if Lcase(strFext)<>"gif" then strFext="gif"
						m_binItem.Position=6
						m_dicForm.Add strInam&"_Width",BinVal2(m_binItem.Read(2))
						m_dicForm.Add strInam&"_Height",BinVal2(m_binItem.Read(2))
					case "image/png"
						if Lcase(strFext)<>"png" then strFext="png"
						m_binItem.Position=18
						m_dicForm.Add strInam&"_Width",Bin2Val(m_binItem.Read(2))
						m_binItem.Read(2)
						m_dicForm.Add strInam&"_Height",Bin2Val(m_binItem.Read(2))
					case "image/bmp"
						if Lcase(strFext)<>"bmp" then strFext="bmp"
						m_binItem.Position=18
						m_dicForm.Add strInam&"_Width",BinVal2(m_binItem.Read(4))
						m_dicForm.Add strInam&"_Height",BinVal2(m_binItem.Read(4))
					case "application/x-shockwave-flash"
						if Lcase(strFext)<>"swf" then strFext="swf"
						m_binItem.Position=0
						if Ascb(m_binItem.Read(1))=70 then
							m_binItem.Position=8
							strTemp = Num2Str(Ascb(m_binItem.Read(1)), 2 ,8)
							intTemp = Str2Num(Left(strTemp, 5), 2)
							strTemp = Mid(strTemp, 6)
							while (Len(strTemp) < intTemp * 4)
								strTemp = strTemp & Num2Str(Ascb(m_binItem.Read(1)), 2 ,8)
							wend
							m_dicForm.Add strInam&"_Width", Int(Abs(Str2Num(Mid(strTemp, intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 1, intTemp), 2)) / 20)
							m_dicForm.Add strInam&"_Height",Int(Abs(Str2Num(Mid(strTemp, 3 * intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 2 * intTemp + 1, intTemp), 2)) / 20)
						end if
					end select

					m_dicForm.Add strInam&"_Ext",strFext
					m_dicForm.Add strInam&"_From",p_start
					if m_AutoSave<>2 then
						intTemp=GetFerr(lngFsiz,strFext)
						m_dicForm.Add strInam&"_Err",intTemp
						if intTemp=0 then
							if m_AutoSave=0 then
								strFnam=GetTimeStr()
								if strFext<>"" then strFnam=strFnam&"."&strFext
							end if
							m_binItem.SaveToFile m_SavePath&strFnam,2
							m_dicForm.Add strInam,strFnam
						end if
					end if
				else
					m_dicForm.Add strInam&"_Err",-1
				end if
			end if
			else
				m_binItem.Position=0
				m_binItem.Type=2
				m_binItem.Charset=m_Charset
				strTemp=m_binItem.ReadText
				if m_dicForm.Exists(strInam) then
					m_dicForm(strInam) = m_dicForm(strInam)&","&strTemp
				else
					strFormItem=strFormItem&strSplit&strInam
					m_dicForm.Add strInam,strTemp
				end if
			end if

			m_binItem.Close()
			p_start = p_end+intSeparator+2
		loop Until p_start+3>lngRequestSize
		FormItem=Split(strFormItem,strSplit)
		FileItem=Split(strFileItem,strSplit)

		Open = lngRequestSize
	End Function

	Private Function GetTimeStr()
		m_lngTime=m_lngTime+1
		GetTimeStr=m_strDate&Right("00000000"&m_lngTime,8)
	End Function

	Private Function GetFerr(lngFsiz,strFext)
		dim intFerr
		intFerr=0
		if lngFsiz>m_MaxSize and m_MaxSize>0 then
			if m_Error=0 or m_Error=2 then m_Error=m_Error+1
			intFerr=intFerr+1
		end if
		if Instr(1,LCase("/"&m_FileType&"/"),LCase("/"&strFext&"/"))=0 and m_FileType<>"" then
			if m_Error<2 then m_Error=m_Error+2
			intFerr=intFerr+2
		end if
		GetFerr=intFerr
	End Function

	Public Function Save(Item,strFnam)
		Save=false
		if m_dicForm.Exists(Item&"_From") then
			dim intFerr,strFext
			strFext=m_dicForm(Item&"_Ext")
			intFerr=GetFerr(m_dicForm(Item&"_Size"),strFext)
			if m_dicForm.Exists(Item&"_Err") then
				if intFerr=0 then
					m_dicForm(Item&"_Err")=0
				end if
			else
				m_dicForm.Add Item&"_Err",intFerr
			end if
			if intFerr<>0 then Exit Function
			if VarType(strFnam)=2 then
				select case strFnam
					case 0:strFnam=GetTimeStr()
						if strFext<>"" then strFnam=strFnam&"."&strFext
					case 1:strFnam=m_dicForm(Item&"_Name")
				end select
			end if
			m_binItem.Type = 1
			m_binItem.Open
			m_binForm.Position = m_dicForm(Item&"_From")
			m_binForm.CopyTo m_binItem,m_dicForm(Item&"_Size")
			m_binItem.SaveToFile m_SavePath&strFnam,2
			m_binItem.Close()
			if m_dicForm.Exists(Item) then
				m_dicForm(Item)=strFnam
			else
				m_dicForm.Add Item,strFnam
			end if
			Save=true
		end if
	End Function

	Public Function GetData(Item)
		GetData=""
		if m_dicForm.Exists(Item&"_From") then
			if GetFerr(m_dicForm(Item&"_Size"),m_dicForm(Item&"_Ext"))<>0 then Exit Function
			m_binForm.Position = m_dicForm(Item&"_From")
			GetData = m_binForm.Read(m_dicForm(Item&"_Size"))
		end if
	End Function

	Public Function Form(Item)
		if m_dicForm.Exists(Item) then
			Form=m_dicForm(Item)
		else
			Form=""
		end if
	End Function

	Private Function BinVal2(bin)
		dim lngValue,i
		lngValue=0
		for i = lenb(bin) to 1 step -1
			lngValue = lngValue *256 + Ascb(midb(bin,i,1))
	Next
		BinVal2=lngValue
	End Function

	Private Function Bin2Val(bin)
		dim lngValue,i
		lngValue=0
		for i = 1 to lenb(bin)
			lngValue = lngValue *256 + Ascb(midb(bin,i,1))
	Next
		Bin2Val=lngValue
	End Function

	Private Function Num2Str(num, base, lens)
		Dim ret,i
		ret = ""
		while(num >= base)
			i   = num Mod base
			ret = i & ret
			num = (num - i) / base
		wend
		Num2Str = Right(String(lens, "0") & num & ret, lens)
	End Function

	Private Function Str2Num(str, base)
		Dim ret, i
		ret = 0
		for i = 1 to Len(str)
			ret = ret * base + Cint(Mid(str, i, 1))
	Next
		Str2Num = ret
	End Function
	Public Function Error2Info(Item)
		Select Case m_dicForm(Item&"_Err")
			case -1:Error2Info = "上传没有开始"
			case 0: Error2Info = "SUCCESS"
			case 1: Error2Info = "文件因大于 "&m_MaxSize&"字节 而未被保存。"
			case 2: Error2Info = "文件因扩展名不符合而未被保存。"
			case 3: Error2Info = "文件因过大且不符合扩展名。"
			case 4: Error2Info = "异常，不存在上传。"
			case 5: Error2Info = "上传已经取消，请检查总上载数据是否小于 "&m_TotalSize&" 。"
		End Select
	End Function

End Class
%>