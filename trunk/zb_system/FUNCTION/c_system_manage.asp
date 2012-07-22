<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)2008-5-30
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:
'// 程序版本:
'// 单元名称:    c_system_manage.asp
'// 开始时间:    2005.02.11
'// 最后修改:
'// 备    注:
'///////////////////////////////////////////////////////////////////////////////

'*********************************************************
' 目的：
'*********************************************************
Function ExportPageBar(PageNow,PageAll,PageLength,Url)

If PageAll=0 Then
	Exit Function
End if

Dim s
Dim i

'Dim PageNow
'Dim PageAll
'Dim PageLength
Dim PageFrist
Dim PageLast
Dim PagePrevious
Dim PageNext
Dim PageBegin
Dim PageEnd

PageFrist = 1
PageLast = PageAll

PageBegin = PageNow
PageEnd = PageBegin + PageLength - 1

If PageEnd > PageAll Then
	PageEnd = PageAll
	PageBegin = PageAll - PageLength + 1
	If PageBegin < 1 Then
		PageBegin = 1
	End If
End If

s=s &"<a href='"&Url & PageFrist &"'>"& "&lt;&lt;" &"</a> "

For i=PageBegin To PageEnd
	If i=PageNow Then
		s=s &"<span>"& Replace(ZC_MSG036,"%s",i) &"</span> "
	Else
		s=s &"<a href='"&Url & i  &"'>"& Replace(ZC_MSG036,"%s",i) &"</a> "
	End If
Next

s=s &"<a href='"&Url & PageLast  &"'>"& "&gt;&gt;" &"</a> "

ExportPageBar=s

End Function



'*********************************************************
' 目的：    Manager Articles
'*********************************************************
Function ExportArticleList(intPage,intCate,intLevel,intTitle)

Call Add_Response_Plugin("Response_Plugin_ArticleMng_SubMenu",MakeSubMenu(ZC_MSG168 & "","../cmd.asp?act=ArticleEdt&amp;webedit=" & ZC_BLOG_WEBEDIT,"m-left",False))

	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim intPageAll

	Call CheckParameter(intPage,"int",1)
	Call CheckParameter(intCate,"int",-1)
	Call CheckParameter(intLevel,"int",-1)
	Call CheckParameter(intTitle,"sql",-1)
	intTitle=vbsunescape(intTitle)
	intTitle=FilterSQL(intTitle)

	Response.Write "<div class=""divHeader"">" & ZC_MSG067 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_ArticleMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"



	Response.Write "<form class=""search"" id=""edit"" method=""post"" action=""../admin/admin.asp?act=ArticleMng"">"

	Response.Write "<p>"&ZC_MSG158&":</p><p>"

	Response.Write ZC_MSG012&" <select class=""edit"" size=""1"" id=""cate"" name=""cate"" style=""width:100px;"" ><option value=""-1"">"&ZC_MSG157&"</option> "

	Dim aryCateInOrder : aryCateInOrder=GetCategoryOrder()
	Dim m,n
	If IsArray(aryCateInOrder) Then
	For m=LBound(aryCateInOrder)+1 To Ubound(aryCateInOrder)
		If Categorys(aryCateInOrder(m)).ParentID=0 Then
			Response.Write "<option value="""&Categorys(aryCateInOrder(m)).ID&""">"&TransferHTML( Categorys(aryCateInOrder(m)).Name,"[html-format]")&"</option>"

			For n=0 To UBound(aryCateInOrder)
				If Categorys(aryCateInOrder(n)).ParentID=Categorys(aryCateInOrder(m)).ID Then
					Response.Write "<option value="""&Categorys(aryCateInOrder(n)).ID&""">└"&TransferHTML( Categorys(aryCateInOrder(n)).Name,"[html-format]")&"</option>"
				End If
			Next
		End If
	Next
	End If
	Response.Write "</select>&nbsp;&nbsp;&nbsp;&nbsp;"

	Response.Write ZC_MSG061&" <select class=""edit"" size=""1"" id=""level"" name=""level"" style=""width:80px;"" ><option value=""-1"">"&ZC_MSG157&"</option> "

	For i=LBound(ZVA_Article_Level_Name)+1 to Ubound(ZVA_Article_Level_Name)
			Response.Write "<option value="""&i&""" "
			Response.Write ">"&ZVA_Article_Level_Name(i)&"</option>"
	Next
	Response.Write "</select>"

	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"&ZC_MSG224&" <input id=""title"" name=""title"" style=""width:250px;"" type=""text"" value="""" /> "
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type=""submit"" class=""button"" value="""&ZC_MSG087&"""/>"

	Response.Write "</p></form>"



	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	strSQL="WHERE ([log_CateID]>0) AND ([log_Level]>0) AND (1=1) "

	If CheckRights("Root")=False Then strSQL= strSQL & "AND [log_AuthorID] = " & BlogUser.ID

	If intCate<>-1 Then
		Dim strSubCateID : strSubCateID=Join(GetSubCateID(intCate,True),",")
		strSQL= strSQL & " AND [log_CateID] IN (" & strSubCateID & ")"
	End If

	If intLevel<>-1 Then
		strSQL= strSQL & " AND [log_Level] = " & intLevel
	End If

	If intTitle<>"-1" Then
		If ZC_MSSQL=False Then
			strSQL = strSQL & "AND ( (InStr(1,LCase([log_Title]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('" & intTitle &"'),0)<>0) )"
		Else
			strSQL = strSQL & "AND ( (CHARINDEX('" & intTitle &"',[log_Title]))<>0) OR (CHARINDEX('" & intTitle &"',[log_Intro])<>0) OR (CHARINDEX('" & intTitle &"',[log_Content])<>0) "
		End If
	End If

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder"">"
	Response.Write "<tr><th width=""5%"">"& ZC_MSG076 &"</th><th width=""14%"">"& ZC_MSG012 &"</th><th width=""14%"">"& ZC_MSG003 &"</th><th>"& ZC_MSG060 &"</th><th width=""14%"">"& ZC_MSG075 &"</th><th width=""14%""></th></tr>"

	objRS.Open("SELECT * FROM [blog_Article] "& strSQL &" ORDER BY [log_PostTime] DESC")
	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			Response.Write "<tr>"

			Response.Write "<td>" & objRS("log_ID") & "</td>"

			Dim Category
			For Each Category in Categorys
				If IsObject(Category) Then
					If Category.ID=objRS("log_CateID") Then
						Response.Write "<td>"
						If Not Category.ParentID=0 Then
							dim objRS2
							Set ObjRS2=objConn.Execute("SELECT [cate_name] FROM [blog_Category] WHERE cate_id="&Category.ParentID&"")
							Response.Write objRS2("cate_Name")
							objRS2.Close
							Set ObjRS2=Nothing
							Response.Write "&nbsp;-->&nbsp;"
						end if
						Response.Write Left(Category.Name,6)
						Response.Write "</td>"
					End If
				End If
			Next

			Dim User
			For Each User in Users
				If IsObject(User) Then
					If User.ID=objRS("log_AuthorID") Then
						Response.Write "<td>" & User.Name & "</td>"
					End If
				End If
			Next

			'Response.Write "<td>" & ZVA_Article_Level_Name(objRS("log_Level")) & "</td>"
			If Len(objRS("log_Title"))>28 Then
				Response.Write "<td><a href=""../view.asp?id=" & objRS("log_ID") & """ title="""& Replace(objRS("log_Title"),"""","") &""" target=""_blank"">" & Left(objRS("log_Title"),14) & "..." & "</a></td>"
			Else
				Response.Write "<td><a href=""../view.asp?id=" & objRS("log_ID") & """ title="""& Replace(objRS("log_Title"),"""","") &""" target=""_blank"">" & objRS("log_Title") & "</a></td>"
			End If
			Response.Write "<td>" & FormatDateTime(objRS("log_PostTime"),vbShortDate) & "</td>"
			Response.Write "<td align=""center""><a href=""../cmd.asp?act=ArticleEdt&amp;webedit="& ZC_BLOG_WEBEDIT &"&amp;id=" & objRS("log_ID") & """><img src=""../image/admin/page_edit.png"" alt=""" & ZC_MSG100 & """ title=""" & ZC_MSG100 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			Response.Write "<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=ArticleDel&amp;id=" & objRS("log_ID") & """><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	Response.Write "</table>"

	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"../admin/admin.asp?act=ArticleMng&amp;cate="&ReQuest("cate")&"&amp;level="&ReQuest("level")&"&amp;title="&Escape(ReQuest("title")) & "&amp;page=")

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	Response.Write "</div>"

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aArticleMng"");</script>"

	objRS.Close
	Set objRS=Nothing

	ExportArticleList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manager SinglePages
'*********************************************************
Function ExportPageList(intPage,intCate,intLevel,intTitle)
'Call SetBlogHint_Custom(ZC_MSG334)
Call Add_Response_Plugin("Response_Plugin_ArticleMng_SubMenu",MakeSubMenu(ZC_MSG328 & "","../cmd.asp?act=ArticleEdt&amp;type=Page&amp;webedit=" & ZC_BLOG_WEBEDIT,"m-left",False))

	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim intPageAll

	Call CheckParameter(intPage,"int",1)
	Call CheckParameter(intCate,"int",-1)
	Call CheckParameter(intLevel,"int",-1)
	Call CheckParameter(intTitle,"sql",-1)
	intTitle=vbsunescape(intTitle)
	intTitle=FilterSQL(intTitle)

	Response.Write "<div class=""divHeader"">" & ZC_MSG327 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_ArticleMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"



	Response.Write "<form class=""search"" id=""edit"" method=""post"" action=""../admin/admin.asp?act=ArticleMng&amp;type=Page"">"

	Response.Write "<p>"&REPLACE(ZC_MSG158,ZC_MSG048,ZC_MSG330)&":</p><p>"

	Response.Write ZC_MSG061&" <select class=""edit"" size=""1"" id=""level"" name=""level"" style=""width:80px;"" ><option value=""-1"">"&ZC_MSG157&"</option> "

	For i=LBound(ZVA_Article_Level_Name)+1 to Ubound(ZVA_Article_Level_Name)
			Response.Write "<option value="""&i&""" "
			Response.Write ">"&Replace(ZVA_Article_Level_Name(i),ZC_MSG048,ZC_MSG330) &"</option>"
	Next
	Response.Write "</select>"

	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"&ZC_MSG224&" <input id=""title"" name=""title"" style=""width:250px;"" type=""text"" value="""" /> "
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type=""submit"" class=""button"" value="""&ZC_MSG087&"""/>"

	Response.Write "</p></form>"



	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	strSQL="WHERE ([log_CateID]=0) AND ([log_Level]>0) AND (1=1) "

	If CheckRights("Root")=False Then strSQL= strSQL & "AND [log_AuthorID] = " & BlogUser.ID

	If intCate<>-1 Then
		Dim strSubCateID : strSubCateID=Join(GetSubCateID(intCate,True),",")
		strSQL= strSQL & " AND [log_CateID] IN (" & strSubCateID & ")"
	End If

	If intLevel<>-1 Then
		strSQL= strSQL & " AND [log_Level] = " & intLevel
	End If

	If intTitle<>"-1" Then
		If ZC_MSSQL=False Then
			strSQL = strSQL & "AND ( (InStr(1,LCase([log_Title]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('" & intTitle &"'),0)<>0) )"
		Else
			strSQL = strSQL & "AND ( (CHARINDEX('" & intTitle &"',[log_Title]))<>0) OR (CHARINDEX('" & intTitle &"',[log_Intro])<>0) OR (CHARINDEX('" & intTitle &"',[log_Content])<>0)"
		End If
	End If

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0""  class=""tableBorder"">"
	Response.Write "<tr><th width='5%'>"& ZC_MSG076 &"</th><th width='14%'>"& ZC_MSG003 &"</th><th>"& ZC_MSG060 &"</th><th width='14%'>"& ZC_MSG075 &"</th><th width='14%'></th></tr>"

	objRS.Open("SELECT * FROM [blog_Article] "& strSQL &" ORDER BY [log_PostTime] DESC")
	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			Response.Write "<tr>"

			Response.Write "<td>" & objRS("log_ID") & "</td>"

			Dim User
			For Each User in Users
				If IsObject(User) Then
					If User.ID=objRS("log_AuthorID") Then
						Response.Write "<td>" & User.Name & "</td>"
					End If
				End If
			Next

			'Response.Write "<td>" & ZVA_Article_Level_Name(objRS("log_Level")) & "</td>"
			If Len(objRS("log_Title"))>28 Then
				Response.Write "<td><a href=""../view.asp?id=" & objRS("log_ID") & """ title="""& Replace(objRS("log_Title"),"""","") &""" target=""_blank"">" & Left(objRS("log_Title"),14) & "..." & "</a></td>"
			Else
				Response.Write "<td><a href=""../view.asp?id=" & objRS("log_ID") & """ title="""& Replace(objRS("log_Title"),"""","") &""" target=""_blank"">" & objRS("log_Title") & "</a></td>"
			End If
			Response.Write "<td>" & FormatDateTime(objRS("log_PostTime"),vbShortDate) & "</td>"
			Response.Write "<td align=""center""><a href=""../cmd.asp?act=ArticleEdt&amp;type=Page&amp;webedit="& ZC_BLOG_WEBEDIT &"&amp;id=" & objRS("log_ID") & """><img src=""../image/admin/page_edit.png"" alt=""" & ZC_MSG100 & """ title=""" & ZC_MSG100 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			Response.Write "<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=ArticleDel&amp;type=Page&amp;id=" & objRS("log_ID") & """><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	Response.Write "</table>"

	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"../admin/admin.asp?act=ArticleMng&amp;type=Page&amp;cate="&ReQuest("cate")&"&amp;level="&ReQuest("level")&"&amp;title="&Escape(ReQuest("title")) & "&amp;page=")

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	Response.Write "</div>"

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aPageMng"");</script>"

	objRS.Close
	Set objRS=Nothing

	ExportPageList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manager Categorys
'*********************************************************
Function ExportCategoryList(intPage)

	Call Add_Response_Plugin("Response_Plugin_CategoryMng_SubMenu",MakeSubMenu(ZC_MSG077 & "","../cmd.asp?act=CategoryEdt","m-left",False))

	Dim i,j

	Response.Write "<div class=""divHeader"">" & ZC_MSG066 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_CategoryMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"


	Call CheckParameter(intPage,"int",1)
'∟

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class='tableBorder'>"
	Response.Write "<tr><th width=""5%""></th><th width=""10%"">"& ZC_MSG076 &"</th><th width=""10%"">"& ZC_MSG079 &"</th><th>"& ZC_MSG001 &"</th><th>"& ZC_MSG147 &"</th><th width=""14%""></th></tr>"

	Dim aryCateInOrder
	aryCateInOrder=GetCategoryOrder()

	If IsArray(aryCateInOrder) Then
	For i=LBound(aryCateInOrder)+1 To Ubound(aryCateInOrder)

		If Categorys(aryCateInOrder(i)).ParentID=0 Then

			Response.Write "<tr><td align=""center""><img width=""16"" src=""../image/admin/folder.png"" alt="""" /></td>"
			Response.Write "<td>" & Categorys(aryCateInOrder(i)).ID & "</td>"
			Response.Write "<td>" & Categorys(aryCateInOrder(i)).Order & "</td>"
			Response.Write "<td><a href=""../catalog.asp?cate="& Categorys(aryCateInOrder(i)).ID &"""  target=""_blank"">" & Categorys(aryCateInOrder(i)).Name & "</a></td>"
			Response.Write "<td>" & Categorys(aryCateInOrder(i)).Alias & "</td>"
			Response.Write "<td align=""center""><a href=""../cmd.asp?act=CategoryEdt&amp;id="& Categorys(aryCateInOrder(i)).ID &"""><img src=""../image/admin/folder_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=CategoryDel&amp;id="& Categorys(aryCateInOrder(i)).ID &"""></a><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></td>"
			Response.Write "</tr>"

			For j=1 To UBound(aryCateInOrder)

				If Categorys(aryCateInOrder(j)).ParentID=Categorys(aryCateInOrder(i)).ID Then
					Response.Write "<tr><td align=""center""><img width=""16"" src=""../image/admin/arrow_turn_right.png"" alt="""" /></td>"
					Response.Write "<td>" & Categorys(aryCateInOrder(j)).ID & "</td>"
					Response.Write "<td>" & Categorys(aryCateInOrder(j)).Order & "</td>"
					Response.Write "<td><a href=""../../catalog.asp?cate="& Categorys(aryCateInOrder(j)).ID &"""  target=""_blank"">" & Categorys(aryCateInOrder(j)).Name & "</a></td>"
					Response.Write "<td>" & Categorys(aryCateInOrder(j)).Alias & "</td>"
					Response.Write "<td align=""center""><a href=""../cmd.asp?act=CategoryEdt&amp;id="& Categorys(aryCateInOrder(j)).ID &"""><img src=""../image/admin/folder_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=CategoryDel&amp;id="& Categorys(aryCateInOrder(j)).ID &"""></a><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></td>"
					Response.Write "</tr>"
				End If

			Next

		End If

	Next
	End If

	Response.Write "</table>"

	Response.Write "</div>"

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aCategoryMng"");</script>"

	ExportCategoryList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manager Comments
'*********************************************************
Function ExportCommentList(intPage,intContent)

	'Call Add_Response_Plugin("Response_Plugin_CommentMng_SubMenu",MakeSubMenu(ZC_MSG211 & "","../cmd.asp?act=CommentEdt","m-left",False))

	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim intPageAll

	Call CheckParameter(intPage,"int",1)
	intContent=FilterSQL(intContent)

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	strSQL=strSQL&" WHERE  ([log_ID]>0) "

	If CheckRights("Root")=False Then strSQL=strSQL & "AND( ([comm_AuthorID] = " & BlogUser.ID & " ) OR ((SELECT [log_AuthorID] FROM [blog_Article] WHERE [blog_Article].[log_ID]=[blog_Comment].[log_ID])=" & BlogUser.ID & " )) "

	If Trim(intContent)<>"" Then strSQL=strSQL & " AND ( ([comm_Author] LIKE '%" & intContent & "%') OR ([comm_Content] LIKE '%" & intContent & "%') OR ([comm_HomePage] LIKE '%" & intContent & "%') ) "

	Response.Write "<div class=""divHeader"">" & ZC_MSG068 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_CommentMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"



	Response.Write "<form class=""search"" id=""edit"" method=""post"" action=""../admin/admin.asp?act=CommentMng"">"
	Response.Write "<p>"&ZC_MSG287&":</p><p>"

	Response.Write " "&ZC_MSG224&" <input id=""intContent"" name=""intContent"" style=""width:250px;"" type=""text"" value="""" /> "
	Response.Write "<input type=""submit"" class=""button"" value="""&ZC_MSG087&"""/>"

	Response.Write "</p></form>"

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder"">"
	Response.Write "<tr><th width=""5%""></th><th width='5%'>"& ZC_MSG076 &"</th><th width=""5%"">"&ZC_MSG331&"</th><th width='14%'>"& ZC_MSG001 &"</th><th>"& ZC_MSG055 &"</th><th width='15%'></th><th width='5%'  align='center'><a href='' onclick='BatchSelectAll();return false'>"& ZC_MSG229 &"</a></th></tr>"'

	objRS.Open("SELECT * FROM [blog_Comment] "& strSQL &" ORDER BY [comm_ID] DESC")
	Dim objArticle
	Set objArticle=New TArticle

	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			objArticle.LoadInfoById objRs("log_ID")

			Response.Write "<tr>"
			Response.Write "<td align=""center""><a href="""&objArticle.URL&"#cmt"&objRS("comm_ID")&""" target=""_blank""><img src=""../image/admin/comment.png"" alt=""" & ZC_MSG212& " @ " & objArticle.HtmlTitle & """ title=""" & ZC_MSG212& " @ " & objArticle.HtmlTitle & """ width=""16"" /></a></td>"
			Response.Write "<td>" & objRS("comm_ID") & "</td>"
			Response.Write "<td>"&IIF(objRs("comm_ParentID")>0,objRs("comm_ParentID"),"")&"</td>"
			If Trim(objRS("comm_Email"))="" Then
			Response.Write "<td>"& objRS("comm_Author") & "</td>"
			Else
			Response.Write "<td><a href=""mailto:"& objRS("comm_Email") &""">" & objRS("comm_Author") & "</a></td>"
			End If

			Response.Write "<td><a id=""mylink"&objRS("comm_ID")&""" href=""$div"&objRS("comm_ID")&"tip?width=400"" class=""betterTip"" title="""&ZC_MSG055&""">" & Left(objRS("comm_Content"),30) & "...</a><div id=""div"&objRS("comm_ID")&"tip"" style=""display:none;""><p>"& objRS("comm_Content") &"</p><br/><p>" & ZC_MSG080 & " : " &objRS("comm_IP") & "</p><p>" & ZC_MSG075 & " : " &objRS("comm_PostTime") & "</p></div></td>"
			Response.Write "<td align=""center""><a href=""../cmd.asp?act=CommentEdt&amp;revid="&objRs("comm_ID")&"&amp;log_id="& objRS("log_ID") &"""><img src=""../image/admin/comments.png"" alt=""" & ZC_MSG333 & """ title=""" & ZC_MSG333 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=""../cmd.asp?act=CommentEdt&amp;amp;id=" & objRS("comm_ID") & "&amp;log_id="& objRS("log_ID") &"&amp;revid="& objRS("comm_ParentID") &"""><img src=""../image/admin/comment_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=""../cmd.asp?act=CommentDel&amp;id=" & objRS("comm_ID") & "&amp;log_id="& objRS("log_ID")  &"&amp;revid="& objRS("comm_ParentID") &""" onclick='return window.confirm("""& ZC_MSG058 &""");'><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
			Response.Write "<td align=""center"" ><input type=""checkbox"" name=""edtDel"" value="""&objRS("comm_ID")&"""/></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next
	Set objArticle=Nothing
	End If

	Response.Write "</table>"

	'For i=1 to objRS.PageCount
	'	strPage=strPage &"<a href='admin.asp?act=CommentMng&amp;page="& i &"'>["& Replace(ZC_MSG036,"%s",i) &"]</a> "
	'Next
	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"admin.asp?act=CommentMng&amp;page=")

	Response.Write "<form id=""frmBatch"" method=""post"" action=""../cmd.asp?act=CommentDelBatch""><input type=""hidden"" id=""edtBatch"" name=""edtBatch"" value=""""/><input class=""button"" type=""submit"" onclick='BatchDeleteAll(""edtBatch"");if(document.getElementById(""edtBatch"").value){return window.confirm("""& ZC_MSG058 &""");}else{return false}' value="""&ZC_MSG228&""" id=""btnPost""/></form>" & vbCrlf

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	Response.Write "</div>"

	objRS.Close
	Set objRS=Nothing

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aCommentMng"");</script>"

	ExportCommentList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manager TrackBacks
'*********************************************************
Function ExportTrackBackList(intPage)


	ExportTrackBackList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manager Users
'*********************************************************
Function ExportUserList(intPage)
	If CheckRights("UserCrt")=True Then
		Call Add_Response_Plugin("Response_Plugin_UserMng_SubMenu",MakeSubMenu(ZC_MSG127 & "","edit_user.asp","m-left",False))
	End If	
	
	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim intPageAll

	Call CheckParameter(intPage,"int",1)

	Response.Write "<div class=""divHeader"">" & ZC_MSG070 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_UserMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"




	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	If CheckRights("Root")=False Then strSQL="WHERE [mem_ID] = " & BlogUser.ID

	objRS.Open("SELECT * FROM [blog_Member] " & strSQL & " ORDER BY [mem_ID] ASC")

	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	If (Not objRS.bof) And (Not objRS.eof) Then

		Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder"">"
		Response.Write "<tr><th width='5%'>"& ZC_MSG076 &"</th><th width='10%'></th><th>"& ZC_MSG001 &"</th><th width='10%'>"& ZC_MSG082 &"</th><th width='10%'>"& ZC_MSG124 &"</th><th width='14%'></th></tr>"

		For i=1 to objRS.PageSize

			Response.Write "<tr>"
			Response.Write "<td>" & objRS("mem_ID") & "</td>"
			Response.Write "<td>" & ZVA_User_Level_Name(objRS("mem_Level")) & "</td>"
			Response.Write "<td><a href=""../../catalog.asp?auth="& objRS("mem_ID") &"""  target=""_blank"">" & objRS("mem_Name") & "</a></td>"

			Response.Write "<td>" & objRS("mem_PostLogs") & "</td>"
			Response.Write "<td>" & objRS("mem_PostComms") & "</td>"

			Response.Write "<td align=""center""><a href=""edit_user.asp?id="& objRS("mem_ID") &"""><img src=""../image/admin/user_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=UserDel&amp;id="& objRS("mem_ID") &"""><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"

			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

		Response.Write "</table>"

	End If

	'For i=1 to objRS.PageCount
	'	strPage=strPage &"<a href='admin.asp?act=UserMng&amp;page="& i &"'>["& Replace(ZC_MSG036,"%s",i) &"]</a> "
	'Next
	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"admin.asp?act=UserMng&amp;page=")

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	Response.Write "</div>"

	objRS.Close
	Set objRS=Nothing

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aUserMng"");</script>"

	ExportUserList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manager Files
'*********************************************************
Function ExportFileList(intPage)

	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim intPageAll

	Call CheckParameter(intPage,"int",1)

	Response.Write "<div class=""divHeader"">" & ZC_MSG071 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_FileMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"



	Response.Write "<form class=""search"" name=""edit"" id=""edit"" method=""post"" enctype=""multipart/form-data"" action=""../cmd.asp?act=FileUpload"">"
	Response.Write "<p>"& ZC_MSG108 &": </p>"
	Response.Write "<p><input type=""file"" id=""edtFileLoad"" name=""edtFileLoad"" size=""40"" />  <input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" name=""B1"" onclick='document.getElementById(""edit"").action=document.getElementById(""edit"").action+""&amp;filename=""+escape(edtFileLoad.value)' /> <input class=""button"" type=""reset"" value="""& ZC_MSG088 &""" name=""B2"" />"
	Response.Write "&nbsp;<input type=""checkbox"" onclick='if(this.checked==true){document.getElementById(""edit"").action=document.getElementById(""edit"").action+""&amp;autoname=1"";}else{document.getElementById(""edit"").action=""../cmd.asp?act=FileUpload"";};SetCookie(""chkAutoFileName"",this.checked,365);' id=""chkAutoName""/><label for=""chkAutoName"">"& ZC_MSG131 &"</label></p></form>"

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	If CheckRights("Root")=False Then strSQL="WHERE [ul_AuthorID] = " & BlogUser.ID

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder"">"
	Response.Write "<tr><th width='5%'>"& ZC_MSG076 &"</th><th width='10%'>"& ZC_MSG003 &"</th><th width=''>"& ZC_MSG001 &"</th><th width='12%'>"& ZC_MSG041 &"</th><th width='12%'>"& ZC_MSG075 &"</th><th width='5%'></th><th width='5%'><a href='' onclick='BatchSelectAll();return false'>"& ZC_MSG229 &"</a></th></tr>"

	objRS.Open("SELECT * FROM [blog_UpLoad] " & strSQL & " ORDER BY [ul_PostTime] DESC")
	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			Response.Write "<tr><td>"&objRS("ul_ID")&"</td>"

			Dim User:For Each User in Users:If IsObject(User) Then:If User.ID=objRS("ul_AuthorID") Then:Response.Write "<td>" & User.Name & "</td>":End If:End If:Next
			If IsNull(objRS("ul_DirByTime"))=False And objRS("ul_DirByTime")<>"" Then
				If CBool(objRS("ul_DirByTime"))=True Then
					Response.Write "<td><a href='../../zb_users/"& ZC_UPLOAD_DIRECTORY &"/"&Year(objRS("ul_PostTime")) & "/" & Month(objRS("ul_PostTime")) & "/"&objRS("ul_FileName")&"' target='_blank'>"&Year(objRS("ul_PostTime")) & "/" & Month(objRS("ul_PostTime")) & "/" &objRS("ul_FileName")&"</a></td>"
				Else
					Response.Write "<td><a href='../../zb_users/"& ZC_UPLOAD_DIRECTORY &"/"&objRS("ul_FileName")&"' target='_blank'>"&objRS("ul_FileName")&"</a></td>"
				End If
			Else
				Response.Write "<td><a href='../../zb_users/"& ZC_UPLOAD_DIRECTORY &"/"&objRS("ul_FileName")&"' target='_blank'>"&objRS("ul_FileName")&"</a></td>"
			End If

			Response.Write "<td>"&objRS("ul_FileSize")&"</td><td>"&FormatDateTime(objRS("ul_PostTime"), 2)&"</td>"
			Response.Write "<td align=""center""><a href='../cmd.asp?act=FileDel&amp;id="&Server.URLEncode(objRS("ul_ID"))&"' onclick='return window.confirm("""& ZC_MSG058 &""");'><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
			Response.Write "<td align=""center"" ><input type=""checkbox"" name=""edtDel"" id=""edtDel"&objRS("ul_ID")&""" value="""&objRS("ul_ID")&"""/></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	Response.Write "</table>"

	Response.Write "<form id=""frmBatch"" method=""post"" action=""../cmd.asp?act=FileDelBatch""><input type=""hidden"" id=""edtBatch"" name=""edtBatch"" value=""""/><input class=""button"" type=""submit"" onclick='BatchDeleteAll(""edtBatch"");if(document.getElementById(""edtBatch"").value){return window.confirm("""& ZC_MSG058 &""");}else{return false}' value="""&ZC_MSG228&""" id=""btnPost""/></form>" & vbCrlf

	'For i=1 to objRS.PageCount
	'	strPage=strPage &"<a href='admin.asp?act=FileMng&amp;page="& i &"'>["& Replace(ZC_MSG036,"%s",i) &"]</a> "
	'Next
	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"admin.asp?act=FileMng&amp;page=")

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	Response.Write "</div><script type=""text/javascript"">if(GetCookie(""chkAutoFileName"")==""true""){document.getElementById(""chkAutoName"").checked=true;document.getElementById(""edit"").action=document.getElementById(""edit"").action+""&amp;autoname=1"";};</script>"
	objRS.Close
	Set objRS=Nothing

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aFileMng"");</script>"

	ExportFileList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manage Setting
'*********************************************************
Function ExportManageList()

	ExportManageList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manager KeyWord
'*********************************************************
Function ExportKeyWordList(intPage)

	ExportKeyWordList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manager Tag
'*********************************************************
Function ExportTagList(intPage)
	Call Add_Response_Plugin("Response_Plugin_TagMng_SubMenu",MakeSubMenu(ZC_MSG136 & "","../cmd.asp?act=TagEdt","m-left",False))

	Dim i
	Dim objRS
	Dim strPage
	Dim intPageAll

	Response.Write "<div class=""divHeader"">" & ZC_MSG141 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_TagMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"





	Call CheckParameter(intPage,"int",1)

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	objRS.Open("SELECT * FROM [blog_Tag] ORDER BY [tag_Name] ASC")

	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder"">"
	Response.Write "<tr><th width=""8%"">"& ZC_MSG076 &"</th><th>"& ZC_MSG001 &"</th><th>"& ZC_MSG016 &"</th><th width=""14%""></th></tr>"

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			Response.Write "<tr>"
			Response.Write "<td>" & objRS("tag_ID") & "</td>"
			Response.Write "<td>" & objRS("tag_Name") & "</td>"
			If IsNull(objRS("tag_Intro"))=True Then
				Response.Write "<td></td>"
			Else
				Response.Write "<td>" & TransferHTML(objRS("tag_Intro"),"[html-format]") & "</td>"
			End If
			Response.Write "<td align=""center""><a href=""../cmd.asp?act=TagEdt&amp;id="& objRS("tag_ID") &"""><img src=""../image/admin/tag_blue_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=TagDel&amp;id="& objRS("tag_ID") &"""><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	Response.Write "</table>"

	'For i=1 to objRS.PageCount
	'	strPage=strPage &"<a href='admin.asp?act=TagMng&amp;page="& i &"'>["& Replace(ZC_MSG036,"%s",i) &"]</a> "
	'Next
	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"admin.asp?act=TagMng&amp;page=")

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	Response.Write "</div>"

	objRS.Close
	Set objRS=Nothing

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aTagMng"");</script>"

	ExportTagList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manager Plugin
'*********************************************************
Function ExportPluginMng()

	On Error Resume Next

	Dim aryPL_Enable()
	Dim aryPL_Disable()

	ReDim aryPL_Enable(0)
	ReDim aryPL_Disable(0)

	Dim aryPL
	aryPL=Split(ZC_USING_PLUGIN_LIST,"|")

	Dim i,j,s,t,m,n

	If ZC_USING_PLUGIN_LIST<>"" Then
		i=UBound(aryPL)
	Else
		i=0
	End If

	ReDim aryPL_Enable(i)


	If Request.QueryString("installed")<>"" Then

		Call InstallPlugin(Request.QueryString("installed"))

	End If

	Dim fso, f, f1, fc
	Set fso = CreateObject("Scripting.FileSystemObject")

	Response.Write "<div class=""divHeader"">" & ZC_MSG107 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_PlugInMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"




	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder"">"
	Response.Write "<tr><th width=""6%"">"& ZC_MSG309 &"</th><th width=""6%"">"& ZC_MSG079 &"</th><th>"& ZC_MSG001 &"</th><th width=""15%"">"& ZC_MSG128 &"</th><th width=""15%"">"& ZC_MSG150 &"</th><th width=""15%"">"& ZC_MSG151 &"</th><th width=""5%""></th><th width=""5%""></th></tr>"

	Dim objXmlFile,strXmlFile



	strXmlFile =BlogPath & "zb_users/theme/" & ZC_BLOG_THEME & "/" & "theme.xml"

	Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
	objXmlFile.async = False
	objXmlFile.ValidateOnParse=False
	objXmlFile.load(strXmlFile)
	If objXmlFile.readyState=4 Then
		If objXmlFile.parseError.errorCode <> 0 Then
		Else

			If CInt(objXmlFile.documentElement.selectSingleNode("plugin/level").text)>0 Then

				If Err.Number=0 Then

					Response.Write "<tr>"
					Response.Write "<td align='center'><img width='16' src='../IMAGE/ADMIN/arrow-3-right.png'/></td>"
					Response.Write "<td>"& "0" &"</td>"
					Response.Write "<td><a id=""mylink"&Left(md5(objXmlFile.documentElement.selectSingleNode("id").text),6)&""" href=""$div"&Left(md5(objXmlFile.documentElement.selectSingleNode("id").text),6)&"tip?width=300"" class=""betterTip"" title=""$content"">" & "" & objXmlFile.documentElement.selectSingleNode("plugin/name").text & "" & "</a><div id=""div"&Left(md5(objXmlFile.documentElement.selectSingleNode("id").text),6)&"tip"" style=""display:none;"">"&objXmlFile.documentElement.selectSingleNode("plugin/note").text&"</div></td>"
					Response.Write "<td>" & "<a target=""_blank"" href=""" & objXmlFile.documentElement.selectSingleNode("author/url").text & """>"& objXmlFile.documentElement.selectSingleNode("author/name").text & "</td>"
					Response.Write "<td>" & objXmlFile.documentElement.selectSingleNode("version").text & "</td>"
					Response.Write "<td>"& objXmlFile.documentElement.selectSingleNode("modified").text &"</td>"
					Response.Write "<td align='center'>"& ZC_MSG311 &"</td>"
					Response.Write "<td align='center'>"
					If BlogUser.Level<=CInt(objXmlFile.documentElement.selectSingleNode("plugin/level").text) Then
						If fso.FileExists(BlogPath & "zb_users/theme/" & ZC_BLOG_THEME & "/plugin/" & objXmlFile.documentElement.selectSingleNode("plugin/path").text) Then
							Response.Write "<a href=""../../ZB_USERS/theme/" & ZC_BLOG_THEME & "/plugin/" & objXmlFile.documentElement.selectSingleNode("plugin/path").text &""">[" & ZC_MSG022 & "]</a>"
						End If
					End If
					Response.Write "</td>"
					Response.Write "</tr>"

				End If

			End If

		End If
	End If
	Set objXmlFile=Nothing

	Set f = fso.GetFolder(BlogPath & "zb_users/plugin/")
	Set fc = f.SubFolders
	For Each f1 in fc

		s=""

		If fso.FileExists(BlogPath & "zb_users/plugin/" & f1.name & "/" & "plugin.xml") Then

			strXmlFile =BlogPath & "zb_users/plugin/" & f1.name & "/" & "plugin.xml"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else
					'If BlogUser.Level<=CInt(objXmlFile.documentElement.selectSingleNode("level").text) Then

			If CheckPluginState(objXmlFile.documentElement.selectSingleNode("id").text) Then
				For j=0 To UBound(aryPL)
					If UCase(aryPL(j))=UCase(objXmlFile.documentElement.selectSingleNode("id").text) Then
						n=j
						Exit For
					End If
				Next
				m=n+1
			Else
				m=""
			End If


			s=s & "<tr>"

			If CheckPluginState(objXmlFile.documentElement.selectSingleNode("id").text) Then
				s=s & "<td align='center'><img width='16' src='../IMAGE/ADMIN/arrow-3-right.png'/></td>"
			Else
				s=s & "<td align='center'><img width='16' src='../IMAGE/ADMIN/MD-stop.png'/></td>"
			End If

			s=s & "<td>"& m &"</td>"
			s=s & "<td><a id=""mylink"&Left(md5(objXmlFile.documentElement.selectSingleNode("id").text),6)&""" href=""$div"&objXmlFile.documentElement.selectSingleNode("id").text&"tip?width=300"" class=""betterTip"" title=""$content"">" & "" & objXmlFile.documentElement.selectSingleNode("name").text & "" & "</a><div id=""div"&objXmlFile.documentElement.selectSingleNode("id").text&"tip"" style=""display:none;"">"&objXmlFile.documentElement.selectSingleNode("note").text&"</div></td>"
			s=s & "<td>" & "<a target=""_blank"" href=""" & objXmlFile.documentElement.selectSingleNode("author/url").text & """>"& objXmlFile.documentElement.selectSingleNode("author/name").text & "</a></td>"
			s=s & "<td>" & objXmlFile.documentElement.selectSingleNode("version").text & "</td>"
			s=s & "<td>"& objXmlFile.documentElement.selectSingleNode("modified").text &"</td>"

				s=s & "<td align='center'>"
			If CheckPluginState(objXmlFile.documentElement.selectSingleNode("id").text) Then
				If CheckRights("PlugInDisable")=True Then
					s=s & "<a href=""../cmd.asp?act=PlugInDisable&amp;name="& Server.URLEncode(objXmlFile.documentElement.selectSingleNode("id").text) &"""><img width='16' title='"&ZC_MSG307&"' alt='"&ZC_MSG307&"' src='../IMAGE/ADMIN/stop.png'/></a>"
				Else

				End If
			Else
				If CheckRights("PlugInActive")=True Then
					s=s & "<a href=""../cmd.asp?act=PlugInActive&amp;name="& Server.URLEncode(objXmlFile.documentElement.selectSingleNode("id").text) &"""><img width='16' title='"&ZC_MSG308&"' alt='"&ZC_MSG308&"' src='../IMAGE/ADMIN/accept.png'/></a>"
				Else
				End If
			End If
			s=s & "</td>"


			s=s & "<td align='center'>"
			If CheckPluginState(objXmlFile.documentElement.selectSingleNode("id").text) Then
				If BlogUser.Level<=CInt(objXmlFile.documentElement.selectSingleNode("level").text) Then
					If fso.FileExists(BlogPath & "zb_users/plugin/" & f1.name & "/" & objXmlFile.documentElement.selectSingleNode("path").text) Then
						s=s & "<a href=""../../ZB_USERS/plugin/" & f1.name & "/" & objXmlFile.documentElement.selectSingleNode("path").text &"""><img width='16' title='"&ZC_MSG022&"' alt='"&ZC_MSG022&"' src='../IMAGE/ADMIN/application_double.png'/></a>"
					End If
				End If
			Else
			End If
			s=s & "</td>"

			s=s & "</tr>"


			If CheckPluginState(objXmlFile.documentElement.selectSingleNode("id").text) Then

				'j=UBound(aryPL_Enable)
				'ReDim Preserve aryPL_Enable(j+1)
				aryPL_Enable(n)=s
			Else
				j=UBound(aryPL_Disable)
				ReDim Preserve aryPL_Disable(j+1)
				aryPL_Disable(j)=s
			End If

				End If
			End If
			Set objXmlFile=Nothing
		End If
	Next

	Response.Write Join(aryPL_Enable)

	Response.Write Join(aryPL_Disable)

	Response.Write "</table>"
	
	Response.Write "</div>"

%>

<%

	Err.Clear

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aPlugInMng"");</script>"

	ExportPluginMng=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function ExportSiteInfo()

	On Error Resume Next

	Dim FoundFso
	FoundFso = False
	FoundFso = IsObjInstalled("Scripting.FileSystemObject")


	Dim objRS
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	Dim allArticle,allCommNums,allTrackBackNums,allViewNums,allUserNums,allCateNums,allTagsNums

	Dim User,i
	For Each User in Users
		If IsObject(User) Then
			Set objRS=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE [log_Level]>1 AND [log_AuthorID]=" & User.ID )
			i=objRS(0)
			objConn.Execute("UPDATE [blog_Member] SET [mem_PostLogs]="&i&" WHERE [mem_ID] =" & User.ID)
			Set objRS=Nothing

			Set objRS=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [comm_AuthorID]=" & User.ID )
			i=objRS(0)
			objConn.Execute("UPDATE [blog_Member] SET [mem_PostComms]="&i&" WHERE [mem_ID] =" & User.ID)
			Set objRS=Nothing
		End If
	Next

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""
	objRS.Open("SELECT COUNT([log_ID])AS allArticle,SUM([log_CommNums]) AS allCommNums,SUM([log_ViewNums]) AS allViewNums,SUM([log_TrackBackNums]) AS allTrackBackNums FROM [blog_Article]")
	If (Not objRS.bof) And (Not objRS.eof) Then
		allArticle=objRS("allArticle")
		allCommNums=objRS("allCommNums")
		allTrackBackNums=objRS("allTrackBackNums")
		allViewNums=objRS("allViewNums")
	End If
	objRS.Close

	objRS.Open("SELECT COUNT([tag_ID])AS allTagsNums FROM [blog_Tag]")
	If (Not objRS.bof) And (Not objRS.eof) Then
		allTagsNums=objRS("allTagsNums")
	End If
	objRS.Close

	objRS.Open("SELECT COUNT([mem_ID])AS allUserNums FROM [blog_Member]")
	If (Not objRS.bof) And (Not objRS.eof) Then
		allUserNums=objRS("allUserNums")
	End If
	objRS.Close

	objRS.Open("SELECT COUNT([cate_ID])AS allCateNums FROM [blog_Category]")
	If (Not objRS.bof) And (Not objRS.eof) Then
		allCateNums=objRS("allCateNums")
	End If
	objRS.Close

	Call CheckParameter(allArticle,"int",0)
	Call CheckParameter(allCommNums,"int",0)
	Call CheckParameter(allTrackBackNums,"int",0)
	Call CheckParameter(allViewNums,"int",0)
	Call CheckParameter(allUserNums,"int",0)
	Call CheckParameter(allCateNums,"int",0)
	Call CheckParameter(allTagsNums,"int",0)

	Response.Write "<div class=""divHeader"">" & ZC_MSG159 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_SiteInfo_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"


	%>

	<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%" class="tableBorder">
	<tr><th height="32" colspan="4"  align="center">&nbsp;<%=ZC_MSG167%></th></tr>
	<tr>
	<td width="20%"><%=ZC_MSG160%></td>
	<td width="30%"><%=BlogUser.Name%> (<%=ZVA_User_Level_Name(BlogUser.Level)%>)</td>
	<td width="20%"><%=ZC_MSG150%></td>
	<td width="30%"><%=ZC_BLOG_VERSION%></td>
	</tr>
	<tr>
	<td width="20%"><%=ZC_MSG082%></td>
	<td width="30%"><%=allArticle%></td>
	<td width="20%"><%=ZC_MSG124%></td>
	<td width="30%"><%=allCommNums%></td>
	</tr>
	<tr>
	<td width="20%"><%=ZC_MSG125%></td>
	<td width="30%"><%=allTrackBackNums%></td>
	<td width="20%"><%=ZC_MSG129%></td>
	<td width="30%"><%=allViewNums%></td>
	</tr>
	<tr>
	<td width="20%"><%=ZC_MSG163%></td>
	<td width="30%"><%=allTagsNums%></td>
	<td width="20%"><%=ZC_MSG162%></td>
	<td width="30%"><%=allCateNums%></td>
	</tr>
	<tr>
	<td width="20%"><%=ZC_MSG306%>/<%=ZC_MSG083%></td>
	<td width="30%"><%=GetNameFormTheme(ZC_BLOG_THEME)%> / <%=ZC_BLOG_CSS%></td>
	<td width="20%"><%=ZC_MSG166%></td>
	<td width="30%"><%=allUserNums%></td>
	</tr>
	<tr>
	<td width="20%">MetaWeblog API</td>
	<td colspan="3" width="80%"><%=GetCurrentHost%>zb_system/xml-rpc/index.asp</td>
	</tr>
<!-- 	<tr>
	<td colspan="4">
	<marquee onmouseover="this.stop()" onmouseout="this.start()"></marquee>
	</td>
	</tr> -->
	</table>
<!--
	<table border="0" cellspacing="0" cellpadding="0" align='center' width="100%" class="tableBorder">
	<tr><th height="32" colspan="4">&nbsp;<%=ZC_MSG164%></th></tr>
	<tr>
	<td width="22%" ><%=ZC_MSG150%></td>
	<td width="27%"><%=ZC_BLOG_VERSION%></td>
	<td width="27%"></td>
	<td width="24%"></td>
	</tr>
	<tr>
	<td width="22%" >FSO </td>
	<td width="27%">
	<%
	If FoundFso Then
		Response.Write "<font color=green><b>ok</b></font>"
	Else
		Response.Write "<font color=red><b>fail</b></font>"
	End If
	%>
	</td>
	<td> Adodb.Stream </td>
	<td><%
	If IsObjInstalled("Adodb.Stream") Then
		Response.Write "<font color=green><b>ok</b></font>"
	Else
		Response.Write "<font color=red><b>fail</b></font>"
	End If
	%>
	</td>
	</tr>
	<tr>
	<td width="22%" >ADODB.Connection</td>
	<td width="27%">
	<%
	If IsObjInstalled("ADODB.Connection") Then
		Response.Write "<font color=green><b>ok</b></font>"
	Else
		Response.Write "<font color=red><b>fail</b></font>"
	End If
	%></td>
	<td> Microsoft.XMLDOM</td>
	<td><%
	If IsObjInstalled("Microsoft.XMLDOM") Then
		Response.Write "<font color=green><b>ok</b></font>"
	Else
		Response.Write "<font color=red><b>fail</b></font>"
	End If
	%>
	</td>
	</tr>
	<tr>
	<td width="22%" >
	MSXML2.ServerXMLHTTP</td>
	<td width="27%">
	<%
	If IsObjInstalled("MSXML2.ServerXMLHTTP") Then
		Response.Write "<font color=green><b>ok</b></font>"
	Else
		Response.Write "<font color=red><b>fail</b></font>"
	End If
	%>
	</td>
	<td > Scripting.Dictionary</td>
	<td><%
	If IsObjInstalled("Scripting.Dictionary") Then
		Response.Write "<font color=green><b>ok</b></font>"
	Else
		Response.Write "<font color=red><b>fail</b></font>"
	End If
	%>
	</td>
	</tr>
	</table>
-->
<%
If Len(ZC_UPDATE_INFO_URL)>0 Then
%>
	<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%" class="tableBorder">
	<tr><th height="32" colspan="4" align="center">&nbsp;<%=ZC_MSG164%>&nbsp;<a href="javascript:updateinfo('?reload');">[<%=ZC_MSG289%>]</a></th></tr>
	<tr><td height="25" colspan="4" id="tdUpdateInfo">
<script language="JavaScript" type="text/javascript">
function updateinfo(s){
	$.post("c_updateinfo.asp"+s,{},
		function(data){
			$("#tdUpdateInfo").html(data);
		}
	)
};

$(document).ready(function(){updateinfo("");});

</script>
	</td></tr>
	</table>
<%
End If
%>
	<br />
<%
	Response.Write "</div>"

	ExportSiteInfo=True

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aSiteInfo"");</script>"

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function ExportFileReBuildAsk()

'Call Add_Response_Plugin("Response_Plugin_AskFileReBuild_SubMenu",MakeSubMenu(ZC_MSG072,"../cmd.asp?act=BlogReBuild","m-left",False))


	Response.Write "<div class=""divHeader2"">" & ZC_MSG073 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_AskFileReBuild_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"

	Response.Write "<iframe frameborder='0' height='500' marginheight='0' marginwidth='0' scrolling='no' width='100%' src='../cmd.asp?act=AskFileReBuild&amp;iframe=1'>"



	Response.Write "</iframe>"

	Response.Write "</div>"

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aAskFileReBuild"");</script>"

	ExportFileReBuildAsk=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function ExportThemeMng()

	On Error Resume Next

	Dim CurrentTheme
	Dim CurrentStyle

	CurrentTheme=ZC_BLOG_THEME
	CurrentStyle=ZC_BLOG_CSS

	Dim Theme_Id
	Dim Theme_Name
	Dim Theme_Url
	Dim Theme_Note
	Dim Theme_Description
	Dim Theme_Pubdate
	Dim Theme_Source_Name
	Dim Theme_Source_Url
	Dim Theme_Author_Name
	Dim Theme_Author_Url
	Dim Theme_ScreenShot
	Dim Theme_Style_Name
	Dim i,j
	Dim aryFileList

	If Request.QueryString("installed")<>"" Then

		Call InstallPlugin(Request.QueryString("installed"))

	End If


	Response.Write "<div class=""divHeader"">" & ZC_MSG291 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_ThemeMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"



	Response.Write "<form id=""frmTheme"" method=""post"" action=""../cmd.asp?act=ThemeSav"">"

	Dim objXmlFile,strXmlFile
	Dim fso, f, f1, fc, s
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "zb_users/theme" & "/")
	Set fc = f.SubFolders
	For Each f1 in fc

		If fso.FileExists(BlogPath & "zb_users/theme" & "/" & f1.name & "/" & "theme.xml") Then

			strXmlFile =BlogPath & "zb_users/theme" & "/" & f1.name & "/" & "theme.xml"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else

					Theme_Id=""
					Theme_Name=""
					Theme_Url=""
					Theme_Note=""
					Theme_Description=""
					Theme_Pubdate=""
					Theme_Source_Name=""
					Theme_Source_Url=""
					Theme_Author_Name=""
					Theme_Author_Url=""
					Theme_ScreenShot=""
					Theme_Style_Name=""

					Theme_Source_Name=objXmlFile.documentElement.selectSingleNode("source/name").text
					Theme_Source_Url=objXmlFile.documentElement.selectSingleNode("source/url").text

					Theme_Author_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					Theme_Author_Url=objXmlFile.documentElement.selectSingleNode("author/url").text

					If Theme_Author_Name="" Then
						Theme_Author_Name=Theme_Source_Name
						Theme_Author_Url=Theme_Source_Url
					End If


					'Theme_Id=f1.name
					Theme_Id=objXmlFile.documentElement.selectSingleNode("id").text
					Theme_Name=objXmlFile.documentElement.selectSingleNode("name").text
					Theme_Url=objXmlFile.documentElement.selectSingleNode("url").text
					Theme_Note=objXmlFile.documentElement.selectSingleNode("note").text
					Theme_Pubdate=objXmlFile.documentElement.selectSingleNode("pubdate").text
					Theme_Description=objXmlFile.documentElement.selectSingleNode("description").text

					Theme_ScreenShot="../../zb_users/theme" &"/" & Theme_Id & "/" & "screenshot.png"




		If UCase(Theme_Id)=UCase(CurrentTheme) Then
			Response.Write "<div class=""theme-now"">"
		Else
			Response.Write "<div class=""theme-other"">"
		End If

		If UCase(Theme_Id) <> UCase(f1.name) Then
			Response.Write "<p style=""color:red;"">ID Error! Should be """& f1.name &"""!!</p>"
		Else
			Response.Write "<p><img width='16' title='' alt='' src='../IMAGE/ADMIN/layout.png'/>&nbsp;&nbsp;ID: <a id=""mylink1"&Left(md5(Theme_Id),6)&""" href=""$div"&Left(md5(Theme_Id),6)&"tip?width=300"" class=""betterTip"" title="""&Theme_Id&""">" & "" & Theme_Id & "" & "</a></p>"
		End If
		Response.Write "<p><a id=""mylink"&Left(md5(Theme_Id),6)&""" href=""$div"&Left(md5(Theme_Id),6)&"tip?width=300"" class=""betterTip"" title="""&Theme_Id&"""><img src=""" & Theme_ScreenShot & """ title=""" & Theme_Name & """ alt=""ScreenShot"" width=""200"" height=""150"" /></a></p>"

		Response.Write "<div id=""div"&Left(md5(Theme_Id),6)&"tip"" style=""display:none;"">"
		Response.Write "<p>"&ZC_MSG001&":" & Theme_Name & "</p>"
		Response.Write "<p>"&ZC_MSG128&":" & Theme_Author_Name & "</p>"
		'Response.Write "<p>"&ZC_MSG054&":" & Theme_Author_Url & "</p>"
		Response.Write "<p>"&ZC_MSG313&":" & Theme_Source_Name & "</p>"
		'Response.Write "<p>"&ZC_MSG054&":" & Theme_Source_Url & "</p>"
		Response.Write "<p>"&ZC_MSG011&":" & Theme_Pubdate & "</p>"
		Response.Write "<p>"&ZC_MSG261&":" & Theme_Modified & "</p>"
		Response.Write "<p>"&ZC_MSG312&":<br />" & TransferHTML(Theme_Description,"[enter]") & "</p>"
		Response.Write "</div>"

		If Theme_Url="" Then
			Response.Write "<p>"&ZC_MSG001&":" & Theme_Name & "</p>"
		Else
			Response.Write "<p>"&ZC_MSG001&":<a target=""_blank"" href=""" & Theme_Url & """>" & Theme_Name & "</a></p>"
		End If

		If Theme_Author_Url="" Then
			Response.Write "<p>"&ZC_MSG128&":" & Theme_Author_Name & "</p>"
		Else
			Response.Write "<p>"&ZC_MSG128&":<a target=""_blank"" href=""" & Theme_Author_Url & """>" & Theme_Author_Name & "</a></p>"
		End If


		Response.Write "<p>"&ZC_MSG011&":" & Theme_Pubdate & "</p>"
		Response.Write "<p>"&ZC_MSG016&":" & Theme_Note & "</p>"
		Response.Write "<p>"&ZC_MSG314&":" & "<select class=""edit"" size=""1"" id=""cate"&Left(md5(Theme_Id),6)&""" name=""cate"&Left(md5(Theme_Id),6)&""" style=""width:120px;"" onchange=""document.getElementById('edtZC_BLOG_THEME').value='"&Theme_Id&"';document.getElementById('edtZC_BLOG_CSS').value=this.options[this.selectedIndex].value""><option value=""""></option>"


		aryFileList=LoadIncludeFiles("zb_users\theme" & "/" & Theme_Id & "/style")

		If IsArray(aryFileList) Then
			j=UBound(aryFileList)
			For i=1 to j
				If (InStr(aryFileList(i),".css")>0) Or (InStr(aryFileList(i),".asp")) Then
					Theme_Style_Name=Replace(aryFileList(i),".css","")
					Theme_Style_Name=Replace(Theme_Style_Name,".asp","")
					If Theme_Id=CurrentTheme And Theme_Style_Name=CurrentStyle Then
						Response.Write " <option selected=""selected"" value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
					Else
						If j=1 Then
							Response.Write " <option selected=""selected"" value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
						ElseIf LCase(Theme_Style_Name)="style" Then
							Response.Write " <option selected=""selected"" value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
						ElseIf LCase(Theme_Style_Name)=LCase(Theme_Id) Then
							Response.Write " <option selected=""selected"" value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
						Else
							If i=1 Then
								Response.Write " <option selected=""selected"" value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
							Else
								Response.Write " <option value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
							End If
						End If
					End If
				End If
			Next
		End If

		Response.Write "</select>"
		Response.Write "&nbsp;&nbsp;<a href='#' onclick='if(!document.getElementById(""cate"&Left(md5(Theme_Id),6)&""").value){return false;}else{document.getElementById(""edtZC_BLOG_THEME"").value="""&Theme_Id&""";document.getElementById(""edtZC_BLOG_CSS"").value=document.getElementById(""cate"&Left(md5(Theme_Id),6)&""").value};$(""#frmTheme"").submit()'><img width='16' title='"&ZC_MSG308&"' alt='"&ZC_MSG308&"' src='../IMAGE/ADMIN/arrow_rotate_anticlockwise.png' /></a></p>"


		Response.Write "</div>"



				End If
			Set objXmlFile=Nothing
			End If

		End If

	Next
	Set fso = nothing

		Response.Write "<input type=""hidden"" name=""edtZC_BLOG_CSS"" id=""edtZC_BLOG_CSS"" value="""" />"
		Response.Write "<input type=""hidden"" name=""edtZC_BLOG_THEME"" id=""edtZC_BLOG_THEME"" value="""" />"


	Response.Write "</form>"
	Response.Write "</div>"
	Err.Clear

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aThemeMng"");</script>"

	ExportThemeMng=True

End Function
'*********************************************************





'*********************************************************
' 目的：    Manager Tag
'*********************************************************
Function ExportFunctionList()

	Dim i,j,s

	SetBlogHint_Custom(Round(Right(10111,3)/111)=1)

	Response.Write "<div class=""divHeader"">" & ZC_MSG343 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_FunctionMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"


	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class='tableBorder'>"
	Response.Write "<tr><th width=""5%""></th><th width=""8%"">"& ZC_MSG079 &"</th><th width=""8%"">"& ZC_MSG076 &"</th><th>"& ZC_MSG001 &"</th><th>"& ZC_MSG147 &"</th><th width=""14%"">"&ZC_MSG061&"</th><th width=""14%"">"&ZC_MSG345&"</th><th width=""14%""></th></tr>"

	Dim aryFunctionInOrder
	aryFunctionInOrder=GetFunctionOrder()

	If IsArray(aryFunctionInOrder) Then
	For i=LBound(aryFunctionInOrder)+1 To Ubound(aryFunctionInOrder)

		s=Functions(aryFunctionInOrder(i)).SidebarID

		If s=0 Then
			s=""
		ElseIf s=1 Then
			s=ZC_MSG344
		ElseIf s>1 Then
			s=ZC_MSG344 & s
		End If

		Response.Write "<tr><td align=""center""><img width=""16"" src=""../image/admin/brick.png"" alt="""" /></td>"
		Response.Write "<td>" & Functions(aryFunctionInOrder(i)).Order & "</td>"
		Response.Write "<td class='funid'>" & Functions(aryFunctionInOrder(i)).ID & "</td>"
		Response.Write "<td>" & Functions(aryFunctionInOrder(i)).Name & "</td>"
		Response.Write "<td>" & Functions(aryFunctionInOrder(i)).HtmlID & "</td>"
		Response.Write "<td>" & Functions(aryFunctionInOrder(i)).Ftype & "</td>"
		Response.Write "<td>" & s & "</td>"
		Response.Write "<td align=""center""><a href=""../cmd.asp?act=CategoryEdt&amp;id="& Functions(aryFunctionInOrder(i)).ID &"""><img src=""../image/admin/brick_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=CategoryDel&amp;id="& Functions(aryFunctionInOrder(i)).ID &"""></a><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></td>"
		Response.Write "</tr>"

	Next
	End If

	Response.Write "</table>"

	Response.Write "<form id=""frmBatch"" method=""post"" action=""""><input type=""hidden"" id=""edtBatch"" name=""edtBatch"" value=""""/><input class=""button"" type=""submit"" onclick='if($(""#edtBatch"").attr(""value"")==""""){return false;}$(""#frmBatch"").attr(""action"",""../cmd.asp?act=FunctionMng"");' value="""&ZC_MSG087&""" id=""btnPost""/>&nbsp;&nbsp;&nbsp;&nbsp;("&ZC_MSG346&")</form>" & vbCrlf

	Response.Write "</div>"

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aFunctionMng"");</script>"

%>
<script type="text/javascript">

function sortFunction(){

	$("#edtBatch").attr('value','');

	$(".funid").each(function(){
	   $("#edtBatch").attr('value',$("#edtBatch").attr('value')+ $(this).html()+'_')
	 });

};

$(document).ready(function(){ 

	$(function() {
		$( ".tableBorder" ).sortable({"items":'tr.color2,tr.color3,tr.color4',stop:function(event, ui){bmx2table();sortFunction();}});
		$( ".tableBorder" ).disableSelection();
	});

});
</script>
<%

	ExportFunctionList=True

End Function
'*********************************************************

%>