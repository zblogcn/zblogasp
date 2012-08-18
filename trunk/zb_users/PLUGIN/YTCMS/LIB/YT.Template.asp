<%
Class YT_Template
	Private aryFileList
	Private themesDir
	Private YTExpressions
	Private Sub Class_Initialize()
		If Not IsObject(objConn) Then Call System_Initialize()
		Set YTExpressions = new YT_Expressions
		themesDir="ZB_USERS/THEME" & "/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY
		aryFileList=LoadIncludeFiles(themesDir)
	End Sub
	Private Sub Class_Terminate()
		If Not IsObject(objConn) Then Call System_Terminate()
		Set YTExpressions = Nothing
	End Sub
	Function List()
		Dim jsonText
		jsonText="["
		If IsArray(aryFileList) Then
			Dim j,i
			j=UBound(aryFileList)
			For i=1 to j
				If i<>1 Then jsonText=jsonText&","
				jsonText=jsonText&Chr(34)&aryFileList(i)&Chr(34)
			Next
		End If
		jsonText=jsonText&"]"
		List = jsonText
	End Function
	Function GetFile(Byval fileName)
		GetFile=LoadFromFile(BlogPath&themesDir&"/"&fileName,"utf-8")
	End Function
	Function YT_Create_CacheFile(DataSource,Template,DataClass)
		dim YTArticle,YT
		dim Text:Text = ""
		dim Texts:Texts = ""
		Dim d,i,c,row()
		Dim RE,Match,Matchs
		Set RE = New RegExp 
			RE.Pattern = "\<\#eval\/(.*?)\#\>" 
			RE.IgnoreCase = True
			RE.Global = True
		If DataClass <> "Cmd" Then
			If IsArray(DataSource) Then
				Call Execute("Set YT = new T"&DataClass)
				For i = LBound(DataSource,2) To UBound(DataSource,2)

					Set d = CreateObject("Scripting.Dictionary")
					If YT.LoadInfoByID(DataSource(0,i)) Then
						Text = Template
						Select Case DataClass
							Case "Article"
								Call d.Add("<#article/id#>",YT.ID)
								Call d.Add("<#article/level#>",YT.Level)
								Call d.Add("<#article/title#>",YT.Title)
								Call d.Add("<#article/intro#>",YT.Intro)
								Call d.Add("<#article/content#>",YT.Content)
								Call d.Add("<#article/posttime#>",YT.PostTime)
								Call d.Add("<#article/commnums#>",YT.Commnums)
								Call d.Add("<#article/viewnums#>",YT.Viewnums)
								Call d.Add("<#article/trackbacknums#>",YT.Trackbacknums)
								Call d.Add("<#article/trackback_url#>",YT.TrackBack)
								Call d.Add("<#article/url#>",YT.Url)

								Call GetCategory()
								Call d.Add("<#article/category/id#>",Categorys(YT.CateID).ID)
								Call d.Add("<#article/category/name#>",Categorys(YT.CateID).HtmlName)
								Call d.Add("<#article/category/order#>",Categorys(YT.CateID).Order)
								Call d.Add("<#article/category/count#>",Categorys(YT.CateID).Count)
								Call d.Add("<#article/category/url#>",Categorys(YT.CateID).HtmlUrl)

								Call GetUser()
								Call d.Add("<#article/author/id#>",Users(YT.AuthorID).ID)
								Call d.Add("<#article/author/name#>",Users(YT.AuthorID).Name)
								Call d.Add("<#article/author/level#>",ZVA_User_Level_Name(Users(YT.AuthorID).Level))
								Call d.Add("<#article/author/email#>",Users(YT.AuthorID).Email)
								Call d.Add("<#article/author/homepage#>",Users(YT.AuthorID).HomePage)
								Call d.Add("<#article/author/count#>",Users(YT.AuthorID).Count)
								Call d.Add("<#article/author/url#>",Users(YT.AuthorID).HtmlUrl)

								Call d.Add("<#article/posttime/longdate#>",FormatDateTime(YT.PostTime,vbLongDate))
								Call d.Add("<#article/posttime/shortdate#>",FormatDateTime(YT.PostTime,vbShortDate))
								Call d.Add("<#article/posttime/longtime#>",FormatDateTime(YT.PostTime,vbLongTime))
								Call d.Add("<#article/posttime/shorttime#>",FormatDateTime(YT.PostTime,vbShortTime))
								Call d.Add("<#article/posttime/year#>",Year(YT.PostTime))
								Call d.Add("<#article/posttime/month#>",Month(YT.PostTime))
								Call d.Add("<#article/posttime/monthname#>",ZVA_Month(Month(YT.PostTime)))
								Call d.Add("<#article/posttime/day#>",Day(YT.PostTime))
								Call d.Add("<#article/posttime/weekday#>",Weekday(YT.PostTime))
								Call d.Add("<#article/posttime/weekdayname#>",ZVA_Week(Weekday(YT.PostTime)))
								Call d.Add("<#article/posttime/hour#>",Hour(YT.PostTime))
								Call d.Add("<#article/posttime/minute#>",Minute(YT.PostTime))
								Call d.Add("<#article/posttime/second#>",Second(YT.PostTime))

								Call d.Add("<#article/commentrss#>",YT.WfwCommentRss)
								Call d.Add("<#article/commentposturl#>",TransferHTML(YT.CommentPostUrl,"[html-format]"))
								Call d.Add("<#article/pretrackback_url#>",TransferHTML(YT.PreTrackBack,"[html-format]"))
								Call d.Add("<#article/trackbackkey#>",YT.TrackBackKey)
								Call d.Add("<#article/commentkey#>",YT.CommentKey)

								Call d.Add("<#article/staticname#>",YT.StaticName)
								'Call d.Add("<#article/category/staticname#>",Categorys(YT.CateID).StaticName)
								'Call d.Add("<#article/author/staticname#>",Users(YT.AuthorID).StaticName)
								Call GetTagsbyTagIDList(YT.Tag)
								Call d.Add("<#article/tagtoname#>",YT.TagToName)

								Call d.Add("<#article/firsttagintro#>",YT.FirstTagIntro)

								Call d.Add("<#article/posttime/monthnameabbr#>",ZVA_Month_Abbr(Month(YT.PostTime)))
								Call d.Add("<#article/posttime/weekdaynameabbr#>",ZVA_Week_Abbr(Weekday(YT.PostTime)))
								
								Call Model(d)
							Case "Comment"
								Call d.Add("<#article/comment/id#>",YT.ID)
								Call d.Add("<#article/comment/count#>",YT.Count)
								Call d.Add("<#article/comment/name#>",YT.Author)
								Call d.Add("<#article/comment/url#>",YT.HomePage)
								Call d.Add("<#article/comment/email#>",YT.SafeEmail)
								Call d.Add("<#article/comment/posttime#>",YT.PostTime)
								Call d.Add("<#article/comment/content#>",YT.Content)
								Call d.Add("<#article/comment/authorid#>",YT.AuthorID)
								Call d.Add("<#article/comment/firstcontact#>",YT.FirstContact)
								Call d.Add("<#article/comment/emailmd5#>",YT.EmailMD5)
								Call d.Add("<#article/comment/urlencoder#>",YT.HomePageForAntiSpam)
								
								Set YTArticle = New TArticle
								If YTArticle.LoadInfoByID(YT.log_ID) Then
									Call d.Add("<#article/id#>",YTArticle.ID)
									Call d.Add("<#article/category/id#>",Categorys(YTArticle.CateID).ID)
									Call d.Add("<#article/title#>",YTArticle.Title)
									Call d.Add("<#article/comment/commenturl#>",YTArticle.Url & "#cmt" & YT.ID)
									Call Model(d)
								End If
								Set YTArticle=Nothing
							Case "Tag"
								Call d.Add("<#article/tag/id#>",YT.ID)
								Call d.Add("<#article/tag/name#>",YT.HtmlName)
								Call d.Add("<#article/tag/intro#>",YT.HtmlIntro)
								Call d.Add("<#article/tag/count#>",YT.Count)
								Call d.Add("<#article/tag/url#>",YT.HtmlUrl)
								Call d.Add("<#article/tag/encodename#>",YT.EncodeName)
						End Select
						Call dToArray(d,row)
						Set Matchs = RE.Execute(Text)   ' 执行搜索。
							For Each Match In Matchs
								If Match.SubMatches(0) <> "" Then
									If Not d.Exists(Match.Value) Then
										Call Execute("Call d.Add(Match.Value,"&Replace(Match.SubMatches(0),"'",Chr(34))&")")
									End If
								End If
							Next
						Set Matchs = Nothing
						Texts = Texts & YTExpressions.YT_Each_Tab(Text,d)
					End If
					Set d = Nothing
				Next
				Set YT = Nothing
			End If
		Else
			'当数据处理类为Cmd时,DataSource为SqlScript
			Dim Rs,Field
			Texts = ""
			Set Rs = objConn.Execute(CStr(DataSource))
				If Not (Rs.EOF and Rs.BOF) Then
					DataSource = Rs.GetRows()
					For i = LBound(DataSource,2) To UBound(DataSource,2)
						c = 0
						Set d = CreateObject("Scripting.Dictionary")
						Text = Template
						For Each Field In Rs.Fields
							Call d.Add("<#field/"& Field.Name &"#>",DataSource(c,i))
							c = c + 1
						Next
						Call dToArray(d,row)
						Set Matchs = RE.Execute(Text)
							For Each Match In Matchs
								If Match.SubMatches(0) <> "" Then
									If Not d.Exists(Match.Value) Then
										Call Execute("Call d.Add(Match.Value,"&Replace(Match.SubMatches(0),"'",Chr(34))&")")
									End If
								End If
							Next
						Set Matchs = Nothing
						Texts = Texts & YTExpressions.YT_Each_Tab(Text,d)
						Set d = Nothing
					Next
				End If
			Set Rs = Nothing
		End If
		Set RE = Nothing
		YT_Create_CacheFile = Texts
	End Function
	Sub dToArray(Byval d,Byref row)
		Dim j
		Dim it
		it = d.Items
		Redim row(d.Count)
		For j = 0 To d.Count - 1
			row(j) = it(j)
		Next
	End Sub
	Function AnalysisTab(TabContent)
		Dim TabCollection,TabCollections()
		Dim DataClass,DataSource
		'获取内容区内的所有YT标签返回数组
		Call YTExpressions.GetTabCollection(TabContent,TabCollections)
		
		Dim T
		For Each TabCollection In TabCollections
			T = TabCollection
			
			If Not IsEmpty(T) Then
				
				'替换内置函数
				Dim RE,Match,Matchs
				Set RE = New RegExp
					RE.Pattern = "\<\#eval\/(.*?)\#\>" 
					RE.IgnoreCase = True
					RE.Global = True
				Set Matchs = RE.Execute(YTExpressions.GeTabContent(T,"$1"))
					For Each Match In Matchs
						If Match.SubMatches(0) <> "" Then
							Call Execute("T = Replace(YTExpressions.GeTabContent(T,""$1""),Match.Value,"&Replace(Match.SubMatches(0),"'",Chr(34))&")")
						End If
					Next
				Set Matchs = Nothing
				Set RE = Nothing
				
				'获取YT处理类,方便以后扩展
				DataClass = GetAttribute(T,"YT")
				DataSource = GetAttribute(T,"DataSource")
				
				If DataClass = "Cmd" Then
					DataSource = CStr(DataSource)
				Else
					Call Execute("DataSource = new YT_"& DataClass & "." & Replace(CStr(DataSource),"'",Chr(34)))
				End If
				TabContent = Replace(TabContent,TabCollection,YT_Create_CacheFile(DataSource,YTExpressions.GeTabContent(TabCollection,"$2"),DataClass))
			End If
		Next
		AnalysisTab = TabContent
	End Function
	Function GetAttribute(TabCollection,AttributeName)
		Dim TabAttribute,TabAttributes
		'获取YT标签内所有属性名称与值
		Set TabAttributes = YTExpressions.GetAttributeCollection(TabCollection)
		'遍历所有属性名称
		For Each TabAttribute in TabAttributes.Keys
			'如果属性存在,取其值
			If UCase(AttributeName) = UCase(TabAttribute) Then
				GetAttribute = TabAttributes(TabAttribute)
				Exit For
			End If
		Next
		Set TabAttributes = Nothing
	End Function
	Sub Model(ByRef d)
		Dim YTModelXML,Node,Object,Field
		Set YTModelXML = new YT_Model_XML
		Set Node = YTModelXML.GetModel(d.Item("<#article/category/id#>"))
			If Not Node Is Nothing Then
				Dim Json
					Json = YT_Data_GetRow(Node.selectSingleNode("Table/Name").Text,d.Item("<#article/id#>"))
					If IsEmpty(Json) Then Exit Sub
					Set Object = jsonToObject(YT_Data_GetRow(Node.selectSingleNode("Table/Name").Text,d.Item("<#article/id#>")))
					For Each Field In Object
						Call d.Add("<#article/model/"&Field.Name&"#>",jsUnEscape(Field.Value))
					Next
					Set Object = Nothing
			End If
		Set Node = Nothing
		Set YTModelXML = Nothing
	End Sub
End Class
%>