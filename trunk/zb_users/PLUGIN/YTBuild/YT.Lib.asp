<%
'///////////////////////////////////////////////////////////////////////////////
'//				Z-Blog
'// 作	 者:    	瑜廷
'// 技术支持:    33195#qq.com
'// 程序名称:    	YT.Build
'// 开始时间:    	2010.12.21
'// 最后修改:    2012.08.24
'// 备	 注:    	only for zblog1.8
'///////////////////////////////////////////////////////////////////////////////
Class YTBuildLib

	Private Url(10)
	Private js
	Private ja
	Private ju
	Private jTag
	Private jtc
	Private pCount
	Private R()
	
	Public Property Get PageCount
		PageCount = pCount
	End Property
	
	Private Sub Class_Initialize()
		Set js = new YTStatic
	End Sub
	
	Private Sub Class_Terminate()
		Set js = Nothing
	End Sub
	
	Function View(Key,C)
		Dim O,j,s
		s = "["
		C = Split(C,",")
		j = 0
		For Each O In C
			Erase R
			Select Case Key
				Case "Cate"
					Call js.GetData(O,Empty,Empty,Empty,pCount,R)
				Case "Auth"
					Call js.GetData(Empty,O,Empty,Empty,pCount,R)
				Case "Tags"
					Call js.GetData(Empty,Empty,Empty,O,pCount,R)
				Case "Date"
					Call js.GetData(Empty,Empty,O,Empty,pCount,R)
			End Select
			If Not IsEmpty(R(0)) Then
				If j <> 0 Then s = s & ","
				s = s & Join(R,",")
				j = j + 1
			End If
		Next
		s = s & "]"
		View = s
	End Function
	
	Function Catalog(Key,C)
		Dim O,j,s
		Dim Tag
		s = "["
		C = Split(C,",")
		j = 0
		For Each O In C
			Erase R
			Select Case Key
				Case "Cate"
					Call js.GetData(O,Empty,Empty,Empty,pCount,R)
					If Not IsEmpty(R(0)) Then R(0) = O
				Case "Auth"
					Call js.GetData(Empty,O,Empty,Empty,pCount,R)
					If Not IsEmpty(R(0)) Then R(0) = O
				Case "Tags"
					Call js.GetData(Empty,Empty,Empty,O,pCount,R)
					R(0)=IIF(IsEmpty(R(0)),Empty,O)
				Case "Date"
					Call js.GetData(Empty,Empty,O,Empty,pCount,R)
					If Not IsEmpty(R(0)) Then R(0) = O
				Case Else
					Call js.GetData(Empty,Empty,Empty,Empty,pCount,R)
					R(0) = 0
			End Select
			If Not IsEmpty(R(0)) Then
				If j <> 0 Then s = s & ","
				s = s & "{"&Chr(34)&"intPageCount"&Chr(34)&":"&pCount&","
				s = s &Chr(34)&"Key"&Chr(34)&":"&Chr(34)&Key&Chr(34)&","
				s = s &Chr(34)&"Type"&Chr(34)&":"&Chr(34)&O&Chr(34)&","
				s = s &Chr(34)&"ID"&Chr(34)&":"&Chr(34)&R(0)&Chr(34)&"}"
				j = j + 1
			End If
		Next
		s = s & "]"
		Catalog = s
	End Function
	
	Function ThreadView(ID)
		Dim UrlRules
		Set ja = new TArticle
			If ja.LoadInfoByID(ID) Then
				If js.View(ID) Then
					If ja.FType=0 Then
						UrlRules=ZC_ARTICLE_REGEX
					Else
						UrlRules=ZC_PAGE_REGEX
					End If
					UrlRules=ParseCustomDirectoryForPath(UrlRules,ZC_STATIC_DIRECTORY,ja.CateID,Empty,Empty,Empty,Empty,ja.ID,ja.StaticName)
					Call CreatDirectoryByCustomDirectory(Replace(Mid(UrlRules,1,InStrRev(UrlRules,"\")),BlogPath,""))
					js.Save(UrlRules)
					ThreadView = True
				Else
					ThreadView = False
				End If
			End If
		Set ja = Nothing
	End Function
	
	Function ThreadCatalog(Key,ID,intIndexPage)
		Dim UrlRules,StaticName
		Select Case Key
			Case "Cate"
				Set jtc = new TCategory
				If jtc.LoadInfoByID(ID) Then
					If js.Catalog(intIndexPage,jtc.ID,Empty,Empty,Empty) Then
						UrlRules=ZC_CATEGORY_REGEX
						Call PageUrl(UrlRules,intIndexPage)
						UrlRules=ParseCustomDirectoryForPath(UrlRules,ZC_STATIC_DIRECTORY,Empty,Empty,Empty,Empty,Empty,jtc.ID,jtc.StaticName)
						Call CreatDirectoryByCustomDirectory(Replace(Mid(UrlRules,1,InStrRev(UrlRules,"\")),BlogPath,""))
						js.Save(UrlRules)
						ThreadCatalog = true
					End If
				End If
				Set jtc = Nothing
			Case "Auth"
				Set ju = new TUser
					If ju.LoadInfoByID(ID) Then
						If js.Catalog(intIndexPage,Empty,ju.ID,Empty,Empty) Then
							UrlRules=ZC_USER_REGEX
							Call PageUrl(UrlRules,intIndexPage)
							UrlRules=ParseCustomDirectoryForPath(UrlRules,ZC_STATIC_DIRECTORY,Empty,Empty,Empty,Empty,Empty,ju.ID,ju.StaticName)
							Call CreatDirectoryByCustomDirectory(Replace(Mid(UrlRules,1,InStrRev(UrlRules,"\")),BlogPath,""))
							js.Save(UrlRules)
							ThreadCatalog = true
						End If
					End If
				Set ju = Nothing
			Case "Tags"
				Set jTag = new TTag
				If jTag.LoadInfoByID(ID) Then
					If js.Catalog(intIndexPage,Empty,Empty,Empty,jTag.Name) Then
						UrlRules=ZC_TAGS_REGEX
						Call PageUrl(UrlRules,intIndexPage)
						UrlRules=ParseCustomDirectoryForPath(UrlRules,ZC_STATIC_DIRECTORY,Empty,Empty,Empty,Empty,Empty,jTag.ID,jTag.EncodeName)
						Call CreatDirectoryByCustomDirectory(Replace(Mid(UrlRules,1,InStrRev(UrlRules,"\")),BlogPath,""))
						js.Save(UrlRules)
						ThreadCatalog = true
					End If
				End If
				Set jTag = Nothing
			Case "Date"
				If js.Catalog(intIndexPage,Empty,Empty,ID,Empty) Then
					UrlRules=ZC_DATE_REGEX
					Call PageUrl(UrlRules,intIndexPage)
					UrlRules=ParseCustomDirectoryForPath(UrlRules,ZC_STATIC_DIRECTORY,Empty,Empty,Year(ID),Month(ID),Empty,Empty,Empty)
					Call CreatDirectoryByCustomDirectory(Replace(Mid(UrlRules,1,InStrRev(UrlRules,"\")),BlogPath,""))
					Dim PrevDate,NextDate,s
					PrevDate=GetTimeData(ID,-1)
					NextDate=GetTimeData(ID,1)
					s="<script>"
					If Not IsEmpty(PrevDate) Then s=s&"$('.month"&Month(ID)&"').find('a').eq(0).attr('href','javascript:void(0)');"
					If Not IsEmpty(NextDate) Then s=s&"$('.month"&Month(ID)&"').find('a').eq(2).attr('href','javascript:void(0)');"
					s=s&"</script>"
					js.html=StaticUrlRules("(\<\/body\>)",js.html,s&vbNewLine&"$1")
					js.Save(UrlRules)
					ThreadCatalog = true
				End If
			Case Else
				If js.Catalog(intIndexPage,Empty,Empty,Empty,Empty) Then
					UrlRules=ZC_DEFAULT_REGEX
					If intIndexPage > 1 Then UrlRules=StaticUrlRules("([a-z]+)(\.[a-z]+)",UrlRules,"$1_"&intIndexPage&"$2")
					UrlRules=ParseCustomDirectoryForPath(UrlRules,Empty,Empty,Empty,Empty,Empty,Empty,Empty,Empty)
					js.Save(UrlRules)
					ThreadCatalog = true
				End If
		End Select
	End Function
	
	Sub PageUrl(Byref Url,Index)
		If Index > 1 Then
			If UCase(Mid(Url,InStrRev(Url,"/"),Len(Url)-InStrRev(Url,"/"))) = "DEFAULT.HTM" Then
				Url=StaticUrlRules("(\/default\.html)",Url,"_"&Index&"$1")
			Else
				Url=StaticUrlRules("(\.html)",Url,"_"&Index&"$1")
			End If
		End If
	End Sub
	
	Function GetTimeData(d,i)
		Dim j,t
		d=Year(d)&"-"&Month(d)
		For Each j In new YTStatic.GetdtmYM
			If IsDate(j) Then
				t=DateAdd("m",i,d)
				t=Year(t)&"-"&Month(t)
				If t=j Then Exit Function
			End If
		Next
		GetTimeData=t
	End Function
	
	Function StaticUrlRules(Pattern,Url,UrlRules)
		Dim RE
		Set RE = New RegExp
			RE.Pattern = Pattern
			RE.IgnoreCase = True
			RE.Global = True
		StaticUrlRules = RE.Replace(Url,UrlRules)
		Set RE = Nothing
	End Function
	
	Function Default()
		Dim l,a
		Dim b,s,f
		'数据量大的时候自动识别是否为BLOG系统，菜鸟模式消耗资源过大，需要生成还是让其手动点击，还是注释掉
		's=LoadFromFile(BlogPath & "ZB_USERS/THEME/" & ZC_BLOG_THEME & "/" & ZC_TEMPLATE_DIRECTORY & "\default.html","utf-8")
		'判断是否为BLOG系统
		'If inStr(s,"<#template:pagebar#>") > 0 Then
		'	Set l=cmd.exec(new YTBuildLib.Catalog(Empty,0))
		'		For Each a In l
		'			For b=1 To a.intPageCount
		'				Default=ThreadCatalog(a.Key,a.ID,b)
		'			Next
		'		Next
		'	Set l=Nothing
		'Else
			b=BlogReBuild_Default
			If b Then
				s=LoadFromFile(BlogPath & "zb_users\cache\default.asp","utf-8")
				If Len(s)>0 Then
					s=Replace(s,"<#ZC_BLOG_HOST#>",ZC_BLOG_HOST)
					js.html=s
					f=ParseCustomDirectoryForPath(ZC_DEFAULT_REGEX,Empty,Empty,Empty,Empty,Empty,Empty,Empty,Empty)
					js.Save(f)
					Default=b
					Exit Function
				End If
			End If
		'End If
		Default=False
	End Function
	
End Class
%>