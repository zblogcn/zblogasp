<%
'注册插件
Call RegisterPlugin("BetterFeed","ActivePlugin_BetterFeed")

'具体的接口挂接
Function ActivePlugin_BetterFeed() 

	'Action_Plugin_ExportRSS_Begin
	Call Add_Action_Plugin("Action_Plugin_ExportRSS_Begin","ExportRSS=BetterFeed():Exit Function")
	
	'管理页加入导航
	'Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu",MakeSubMenu("RSS 优化选项","../zb_users/plugin/BetterFeed/main.asp","m-left",False))
		
End Function


'启用插件
Function InstallPlugin_BetterFeed()

	Call BetterFeed_Initialize
	Call ExportRSS
	
End Function


'配置变量
Dim BetterFeed_Copyright_message
Dim BetterFeed_Addreadmoreinfeed
Dim BetterFeed_Readmore_message
Dim BetterFeed_Addcommentinfeed
Dim BetterFeed_Comment_message
Dim BetterFeed_Commentinfeed
Dim BetterFeed_Commentinfeed_limit
Dim BetterFeed_Commentinfeed_before
Dim BetterFeed_Commentinfeed_layout
Dim BetterFeed_Commentinfeed_after
Dim BetterFeed_Relatedpostinfeed
Dim BetterFeed_Relatedpostinfeed_limit
Dim BetterFeed_Relatedpostinfeed_before
Dim BetterFeed_Relatedpostinfeed_layout
Dim BetterFeed_Relatedpostinfeed_after
Dim BetterFeed_Relatedpostinfeed_sub
Dim BetterFeed_Otherinfeed


'初始化配置
Function BetterFeed_Initialize()
	Dim c
	Set c = New TConfig
	c.Load("BetterFeed")
	If c.Exists("BetterFeed_Copyright_message")=False Then
		c.Write "BetterFeed_Copyright_message","<p>Copyright © 2012</p>"
		c.Write "BetterFeed_Addreadmoreinfeed","True"
		c.Write "BetterFeed_Readmore_message","<p><a href=""%permalink%"" target=""_blank"">继续阅读《%posttitle%》的全文内容...</a></p>"
		c.Write "BetterFeed_Addcommentinfeed","True"
		c.Write "BetterFeed_Comment_message","<p>分类: %category% | Tags: %tags% | <a href=""%permalink%#comment"" target=""_blank"">添加评论</a>(%commentcount%)</p>"
		c.Write "BetterFeed_Commentinfeed","False"
		c.Write "BetterFeed_Commentinfeed_limit","5"
		c.Write "BetterFeed_Commentinfeed_before","<hr /> <h3>最新评论:</h3><ul>"
		c.Write "BetterFeed_Commentinfeed_layout","<li><a href=""%permalink%#cmt%commentid%"">%date% %time%</a>，%authorlink% ： %comment%</li>"
		c.Write "BetterFeed_Commentinfeed_after","</ul>"
		c.Write "BetterFeed_Relatedpostinfeed","True"
		c.Write "BetterFeed_Relatedpostinfeed_limit","5"
		c.Write "BetterFeed_Relatedpostinfeed_before","<h3>相关文章:</h3><ul>"
		c.Write "BetterFeed_Relatedpostinfeed_layout","<li><a href=""%article_url%"">%article_title%</a> (%article_time%)  </li>"
		c.Write "BetterFeed_Relatedpostinfeed_after","</ul>"
		c.Write "BetterFeed_Relatedpostinfeed_sub","<p><a href=""%permalink%#comment"" target=""_blank"">还没有相关文章，您来说两句？</a></p>"
		c.Write "BetterFeed_Otherinfeed",""
		c.Save
		Call SetBlogHint_Custom("第一次安装RSS优化插件，已经为您导入初始配置。")
	End If
	Set c=Nothing
End Function


'配置读取
Function BetterFeed_Config()
	Dim c
	Set c = New TConfig
	c.Load("BetterFeed")
	BetterFeed_Copyright_message = c.Read ("BetterFeed_Copyright_message")
	BetterFeed_Addreadmoreinfeed=c.Read ("BetterFeed_Addreadmoreinfeed")
	BetterFeed_Readmore_message=c.Read ("BetterFeed_Readmore_message")
	BetterFeed_Addcommentinfeed=c.Read ("BetterFeed_Addcommentinfeed")
	BetterFeed_Comment_message=c.Read ("BetterFeed_Comment_message")

	BetterFeed_Commentinfeed=c.Read ("BetterFeed_Commentinfeed")
	BetterFeed_Commentinfeed_limit=c.Read ("BetterFeed_Commentinfeed_limit")
	BetterFeed_Commentinfeed_before=c.Read ("BetterFeed_Commentinfeed_before")
	BetterFeed_Commentinfeed_layout=c.Read ("BetterFeed_Commentinfeed_layout")
	BetterFeed_Commentinfeed_after=c.Read ("BetterFeed_Commentinfeed_after")

	BetterFeed_Relatedpostinfeed=c.Read("BetterFeed_Relatedpostinfeed")
	BetterFeed_Relatedpostinfeed_limit=c.Read("BetterFeed_Relatedpostinfeed_limit")
	BetterFeed_Relatedpostinfeed_before=c.Read("BetterFeed_Relatedpostinfeed_before")
	BetterFeed_Relatedpostinfeed_layout=c.Read("BetterFeed_Relatedpostinfeed_layout")
	BetterFeed_Relatedpostinfeed_after=c.Read("BetterFeed_Relatedpostinfeed_after")
	BetterFeed_Relatedpostinfeed_sub=c.Read("BetterFeed_Relatedpostinfeed_sub")

	BetterFeed_Otherinfeed=c.Read("BetterFeed_Otherinfeed")

	Set c=Nothing
End Function


'*********************************************************
' 目的：重写RssExport
'*********************************************************
Function BetterFeed()
	Dim Rss2Export
	Dim objArticle

	Set Rss2Export = New TNewRss2Export

	With Rss2Export

		.TimeZone=ZC_TIME_ZONE

		.AddChannelAttribute "title",TransferHTML(ZC_BLOG_TITLE,"[html-format]")
		.AddChannelAttribute "link",TransferHTML(BlogHost,"[html-format]")
		.AddChannelAttribute "description",TransferHTML(ZC_BLOG_SUBTITLE,"[html-format]")
		.AddChannelAttribute "generator","RainbowSoft Studio Z-Blog " & ZC_BLOG_VERSION
		.AddChannelAttribute "language",ZC_BLOG_LANGUAGE
		.AddChannelAttribute "pubDate",GetTime(Now())

			Dim i
			Dim objRS
			Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_ID]>0) AND ([log_Level]>2) ORDER BY [log_PostTime] DESC")

			If (Not objRS.bof) And (Not objRS.eof) Then
				'开始准备输出，读取配置
				Call BetterFeed_Config

				For i=1 to ZC_RSS2_COUNT
					Set objArticle=New TArticle
					If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then

						If ZC_RSS_EXPORT_WHOLE Then
						.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).FirstName & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlContent&BetterFeedExport(objArticle),Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
						Else
						.AddItem objArticle.HtmlTitle,Users(objArticle.AuthorID).Email & " (" & Users(objArticle.AuthorID).FirstName & ")",objArticle.HtmlUrl,objArticle.PostTime,objArticle.HtmlUrl,objArticle.HtmlIntro&BetterFeedExport(objArticle),Categorys(objArticle.CateID).HtmlName,objArticle.CommentUrl,objArticle.WfwComment,objArticle.WfwCommentRss,objArticle.TrackBackUrl
						End If

					End If
					objRS.MoveNext
					If objRS.eof Then Exit For
					Set objArticle=Nothing
				Next
			End If

	End With

	Rss2Export.SaveToFile(BlogPath & "zb_users\cache\rss.xml")

	Set Rss2Export = Nothing

	objRS.close
	Set objRS=Nothing
	BetterFeed=True

End Function

'*********************************************************
' 目的：RSS优化输出
'*********************************************************
Function BetterFeedExport(objTArticle)
	
	Dim objArticle
	Set objArticle=objTArticle
	
	Dim intID,strTag,intCate,strTitle,strURL,intCommNums
	
	intID=objArticle.ID
	strTag=objArticle.Tag
	intCate=objArticle.CateID
	strTitle=objArticle.Title
	strURL=objArticle.HtmlUrl
	intCommNums=objArticle.CommNums
	
	Dim strBetterFeedExport
	
	Dim strCopyright_message
	Dim blnAddreadmoreinfeed
	Dim strReadmore_message
	Dim blnAddcommentinfeed
	Dim strComment_message
	Dim blnCommentinfeed
	Dim blnRelatedpostinfeed
	Dim strOtherinfeed


	strCopyright_message = BetterFeed_Copyright_message
	blnAddreadmoreinfeed = BetterFeed_Addreadmoreinfeed
	strReadmore_message = BetterFeed_Readmore_message
	blnAddcommentinfeed = BetterFeed_Addcommentinfeed
	strComment_message = BetterFeed_Comment_message
	blnCommentinfeed = BetterFeed_Commentinfeed
	blnRelatedpostinfeed = BetterFeed_Relatedpostinfeed
	strOtherinfeed = BetterFeed_Otherinfeed


	'版权声明
	If strCopyright_message<>"" Then strBetterFeedExport=strBetterFeedExport & Replace(strCopyright_message,"%blogtitle%",ZC_BLOG_TITLE)
	
	'ReadMore
	If blnAddreadmoreinfeed Then
		strBetterFeedExport=strBetterFeedExport & strReadmore_message
	End If
	
	'显示分类等
	'Tag
	Dim Tag
	Dim t,i,s,j
	If objArticle.Tag<>"" Then
			s=Replace(objArticle.Tag,"}","")
			t=Split(s,"{")
			For i=LBound(t) To UBound(t)
				If t(i)<>"" Then
					If IsObject(Tags(t(i)))=True Then
						j=GetTemplate("TEMPLATE_B_ARTICLE_TAG")
						Tag=Tag & Tags(t(i)).MakeTemplate(j)
					End If
				End If
			Next
	End If

	If blnAddcommentinfeed Then

		strComment_message=Replace(strComment_message,"%commentcount%",intCommNums)
		strComment_message=Replace(strComment_message,"%category%","<a href="""&Categorys(intCate).HtmlUrl&""">"&Categorys(intCate).HtmlName&"</a>")
		if objArticle.Export_Tag() Then
			strComment_message=Replace(strComment_message,"%tags%",Tag)
		else 
			strComment_message=Replace(strComment_message,"%tags%","")
		end if 
		strBetterFeedExport=strBetterFeedExport & strComment_message
	End If

	'相关文章
	If 	blnRelatedpostinfeed Then	
		strBetterFeedExport=strBetterFeedExport & BetterFeedRelateList(intID,strTag)
	End If	
	
	'本文评论
	If 	blnCommentinfeed Then	
		strBetterFeedExport=strBetterFeedExport & BetterFeedCommetList(intID,intCommNums)
	End If
	
	'其它内容
	If strOtherinfeed<>"" Then strBetterFeedExport=strBetterFeedExport & strOtherinfeed
	
	Set objArticle=Nothing
	
	BetterFeedExport=Replace(Replace(Replace(strBetterFeedExport,"%permalink%",strURL),"%posttitle%",strTitle),"%bloglink%",ZC_BLOG_HOST)
	
End Function


'*********************************************************
' 目的：评论列表 
'*********************************************************
Function BetterFeedCommetList(intID,intCommNums)
	

	If intCommNums > 0 Then
	
			Dim intCommentinfeed_limit
			Dim strCommentinfeed_before
			Dim strCommentinfeed_layout
			Dim strCommentinfeed_after

			intCommentinfeed_limit = BetterFeed_Commentinfeed_limit
			strCommentinfeed_before = BetterFeed_Commentinfeed_before
			strCommentinfeed_layout = BetterFeed_Commentinfeed_layout
			strCommentinfeed_after = BetterFeed_Commentinfeed_after

			Dim strC_Count,strC,strT_Count,strT

			Dim objComment,strComment
			Dim objTrackBack

			Dim i
			
			Dim User,Cmt_FirstName,p_FirstName

			Dim objRS
			Set objRS=Server.CreateObject("ADODB.Recordset")
			objRS.CursorType = adOpenKeyset
			objRS.LockType = adLockReadOnly
			objRS.ActiveConnection=objConn
			objRS.Source="SELECT [comm_ID],[log_ID],[comm_AuthorID],[comm_Author],[comm_Content],[comm_Email],[comm_HomePage],[comm_PostTime],[comm_IP],[comm_Agent] FROM [blog_Comment] WHERE [blog_Comment].[log_ID]=" & intID &" UNION ALL SELECT [tb_ID],[log_ID],'',[tb_Title],[tb_Excerpt],[tb_Blog],[tb_URL],[tb_PostTime],[tb_IP],[tb_Agent] from [blog_TrackBack] WHERE [blog_TrackBack].[log_ID]="& intID & " ORDER BY [comm_ID],[comm_PostTime]"

			objRS.Open()

			If (not objRS.bof) And (not objRS.eof) Then

				ReDim aryArticleExportMsgTB(objRS.RecordCount)

				For i=1 To intCommentinfeed_limit

					If IsNumeric(objRS("comm_AuthorID")) Then

						Set objComment=New TComment

						objComment.LoadInfoByID(objRS("comm_ID"))
						
						Call GetUser
						Cmt_FirstName = objComment.Author
						For Each User in Users
							If IsObject(User) Then
								If User.ID=objComment.AuthorID Then
									Cmt_FirstName = User.FirstName
									Exit For 
								End If
							End If
						Next

						'添加回复标签
						If objComment.ParentID<>0 Then 
							Dim objRevComment
							Set objRevComment=New TComment
							objRevComment.LoadInfoByID(objComment.ParentID)
							p_FirstName = objRevComment.Author
							For Each User in Users
								If IsObject(User) Then
									If User.ID=objRevComment.AuthorID Then
										p_FirstName = User.FirstName
										Exit For 
									End If
								End If
							Next
							Set objRevComment=Nothing
						End If 


						strC_Count=strC_Count+1

						strC=strCommentinfeed_layout
						strC=Replace(strC,"%commentid%",objComment.ID)
						If objComment.HomePage<>"" Then
							strC=Replace(strC,"%authorlink%","<a href="""&objComment.HomePage&""">"&Cmt_FirstName&"</a>")
						Else
							strC=Replace(strC,"%authorlink%",Cmt_FirstName)
						End If
						strC=Replace(strC,"%date%",Year(objComment.PostTime) & "-" & Month(objComment.PostTime) & "-" & Day(objComment.PostTime))
						strC=Replace(strC,"%time%",Hour(objComment.PostTime) & ":" & Minute(objComment.PostTime) & ":" & Second(objComment.PostTime))
						strC=Replace(strC,"%comment%",objComment.HtmlContent)

						If objComment.ParentID<>0 Then 
							strC=Replace(strC,"%revlink%","<a href=""%permalink%#cmt"&objComment.ParentID&""">@"&p_FirstName&"</a>")
						Else
							strC=Replace(strC,"%revlink%","")
						End If 

						objComment.Count=strC_Count

						strComment=strC & strComment

						Set objComment=Nothing


					End If

					objRS.MoveNext
					If objRS.eof Then Exit For
				Next

			End if

			objRS.Close()
			Set objRS=Nothing

		End If
		
		BetterFeedCommetList=strCommentinfeed_before & strComment &strCommentinfeed_after
		
End Function


'*********************************************************
' 目的：相关文章列表
'*********************************************************
Function BetterFeedRelateList(intID,strTag)
	If (intID=0) Then Exit Function
	If strTag<>"" Then

	Dim strCC_Count,strCC_ID,strCC_Name,strCC_Url,strCC_PostTime,strCC_Title
	Dim strCC
	Dim i
	Dim j
	Dim objRS
	Dim strSQL

	Dim intRelatedpostinfeed_limit
	Dim strRelatedpostinfeed_before
	Dim strRelatedpostinfeed_after
	Dim strRelatedpostinfeed_layout

	intRelatedpostinfeed_limit = BetterFeed_Relatedpostinfeed_limit
	strRelatedpostinfeed_before = BetterFeed_Relatedpostinfeed_before
	strRelatedpostinfeed_after = BetterFeed_Relatedpostinfeed_after
	strRelatedpostinfeed_layout = BetterFeed_Relatedpostinfeed_layout


	Dim strOutput
	strOutput=""

	Set objRS=Server.CreateObject("ADODB.Recordset")

	strSQL="SELECT TOP "& ZC_MUTUALITY_COUNT &" [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_Level]>2) AND [log_ID]<>"& intID
	strSQL = strSQL & " AND ("

	Dim aryTAGs
	strTag=Replace(strTag,"}","")
	aryTAGs=Split(strTag,"{")

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
		For i=1 To intRelatedpostinfeed_limit '相关文章数目
		Dim objArticle
		Set objArticle=New TArticle
		If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then

		    strCC_Count=strCC_Count+1
		    strCC_ID=objArticle.ID
		    strCC_Url=objArticle.Url
		    strCC_PostTime=objArticle.PostTime
		    strCC_Title=objArticle.Title

'		    Application.Lock
'		    strCC=Application(ZC_BLOG_CLSID & "TEMPLATE_B_ARTICLE_Mutuality")
'		    Application.UnLock
'
'		    strCC=Replace(strCC,"<#article/mutuality/id#>",strCC_ID)
'		    strCC=Replace(strCC,"<#article/mutuality/url#>",strCC_Url)
'		    strCC=Replace(strCC,"<#article/mutuality/posttime#>",strCC_PostTime)
'		    strCC=Replace(strCC,"<#article/mutuality/name#>",strCC_Title)

			strCC=strRelatedpostinfeed_layout

		    strCC=Replace(strCC,"%article_id%",strCC_ID)    
		    strCC=Replace(strCC,"%article_url%",strCC_Url)  
		    strCC=Replace(strCC,"%article_time%",strCC_PostTime) 
		    strCC=Replace(strCC,"%article_title%",strCC_Title) 
			
			strOutput=strOutput & strCC

		End If
		objRS.MoveNext
		If objRS.eof Then Exit For
		Set objArticle=nothing
		Next

	End if

	objRS.Close()
	Set objRS=Nothing
	End If

	If strOutput<>"" then
		BetterFeedRelateList= strRelatedpostinfeed_before + strOutput + strRelatedpostinfeed_after
	Else
		BetterFeedRelateList = BetterFeed_Relatedpostinfeed_sub
	End If
	
End Function

%>