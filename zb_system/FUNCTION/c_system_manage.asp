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
Function ExportArticleList(intPage,intCate,intLevel,bolIstop,intTitle)

'Call Add_Response_Plugin("Response_Plugin_ArticleMng_SubMenu",MakeSubMenu(ZC_MSG168 & "","../cmd.asp?act=ArticleEdt&amp;webedit=" & ZC_BLOG_WEBEDIT,"m-left",False))

	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim intPageAll

	Call CheckParameter(intPage,"int",1)
	Call CheckParameter(intCate,"int",-1)
	Call CheckParameter(intLevel,"int",-1)
	Call CheckParameter(bolIstop,"bool",False)
	Call CheckParameter(intTitle,"sql",-1)
	intTitle=vbsunescape(intTitle)
	intTitle=FilterSQL(intTitle)

	Response.Write "<div class=""divHeader"">" & ZC_MSG067 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_ArticleMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"



	Response.Write "<form class=""search"" id=""edit"" method=""post"" action=""../admin/admin.asp?act=ArticleMng"">"

	Response.Write "<p>"&ZC_MSG158&":&nbsp;&nbsp;"

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

	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<label><input type=""checkbox"" name=""istop"" id=""istop"" value=""True""/>&nbsp;"&ZC_MSG051&"</label>"

	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input id=""title"" name=""title"" style=""width:250px;"" type=""text"" value="""" /> "
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type=""submit"" class=""button"" value="""&ZC_MSG087&"""/>"

	Response.Write "</p></form>"

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	strSQL="WHERE ([log_Type]=0) AND ([log_Level]>0) AND (1=1) "

	If CheckRights("Root")=False And CheckRights("ArticleAll")=False Then strSQL= strSQL & "AND [log_AuthorID] = " & BlogUser.ID

	If intCate<>-1 Then
		Dim strSubCateID : strSubCateID=Join(GetSubCateID(intCate,True),",")
		strSQL= strSQL & " AND [log_CateID] IN (" & strSubCateID & ")"
	End If

	If intLevel<>-1 Then
		strSQL= strSQL & " AND [log_Level] = " & intLevel
	End If

	If bolIstop=True Then
		strSQL= strSQL & " AND [log_IsTop] <> 0"
	End If

	If intTitle<>"-1" Then
		If ZC_MSSQL_ENABLE=False Then
			strSQL = strSQL & "AND ( (InStr(1,LCase([log_Title]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('" & intTitle &"'),0)<>0) )"
		Else
			strSQL = strSQL & "AND ( (CHARINDEX('" & intTitle &"',[log_Title])<>0) OR (CHARINDEX('" & intTitle &"',[log_Intro])<>0) OR (CHARINDEX('" & intTitle &"',[log_Content])<>0) )"
		End If
	End If

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder tableBorder-thcenter"">"
	Response.Write "<tr><th width=""5%"">"& ZC_MSG076 &"</th><th width=""14%"">"& ZC_MSG012 &"</th><th width=""12%"">"& ZC_MSG003 &"</th><th>"& ZC_MSG060 &"</th><th width=""14%"">"& ZC_MSG075 &"</th><th width=""6%"">"& ZC_MSG013 &"</th><th width=""9%"">"& ZC_MSG061 &"</th><th width=""12%""></th></tr>"

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
						end if
						Response.Write Category.Name
						Response.Write "</td>"
					End If
				End If
			Next

			Call GetUsersbyUserIDList(objRS("log_AuthorID"))
			Dim User
			For Each User in Users
				If IsObject(User) Then
					If User.ID=objRS("log_AuthorID") Then
						Response.Write "<td>" & User.Name & "</td>"
					End If
				End If
			Next

			Response.Write "<td><div style='overflow:hidden;height:1.5em;'><a href="""&IIf(objRs("log_Level")=1,"../cmd.asp?act=ArticleEdt&amp;webedit="& ZC_BLOG_WEBEDIT &"&amp;id=" & objRS("log_ID"),"../../view.asp?nav=" & objRS("log_ID")) & """ title="""& Replace(objRS("log_Title"),"""","") &""" target=""_blank"">" & objRS("log_Title") & "</a></div></td>"
			Response.Write "<td>" & FormatDateTime(objRS("log_PostTime"),vbShortDate) & "</td>"
			Response.Write "<td>" & objRS("log_CommNums") & "</td>"
			Response.Write "<td>" & ZVA_Article_Level_Name(objRS("log_Level")) & "</td>"
			Response.Write "<td align=""center""><a href=""../cmd.asp?act=ArticleEdt&amp;webedit="& ZC_BLOG_WEBEDIT &"&amp;id=" & objRS("log_ID") & """><img src=""../image/admin/page_edit.png"" alt=""" & ZC_MSG100 & """ title=""" & ZC_MSG100 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			Response.Write "<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=ArticleDel&amp;id=" & objRS("log_ID") & """><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	Response.Write "</table>"

	If  intPageAll>1 Then 
		strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"../admin/admin.asp?act=ArticleMng&amp;cate="&ReQuest("cate")&"&amp;level="&ReQuest("level")&"&amp;title="&Escape(ReQuest("title")) & "&amp;page=")

		Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	End If 

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

Call Add_Response_Plugin("Response_Plugin_ArticleMng_SubMenu",MakeSubMenu(ZC_MSG113 & "","../cmd.asp?act=ArticleEdt&amp;type=Page&amp;webedit=" & ZC_BLOG_WEBEDIT,"m-left",False))

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

	Response.Write "<div class=""divHeader"">" & ZC_MSG111 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_ArticleMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"



	Response.Write "<form class=""search"" id=""edit"" method=""post"" action=""../admin/admin.asp?act=ArticleMng&amp;type=Page"">"

	Response.Write "<p>"&REPLACE(ZC_MSG158,ZC_MSG048,ZC_MSG160)&":&nbsp;&nbsp;&nbsp;&nbsp;"

	Response.Write ZC_MSG061&" <select class=""edit"" size=""1"" id=""level"" name=""level"" style=""width:80px;"" ><option value=""-1"">"&ZC_MSG157&"</option> "

	For i=LBound(ZVA_Article_Level_Name)+1 to Ubound(ZVA_Article_Level_Name)
			Response.Write "<option value="""&i&""" "
			Response.Write ">"&Replace(ZVA_Article_Level_Name(i),ZC_MSG048,ZC_MSG160) &"</option>"
	Next
	Response.Write "</select>"

	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input id=""title"" name=""title"" style=""width:250px;"" type=""text"" value="""" /> "
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type=""submit"" class=""button"" value="""&ZC_MSG087&"""/>"

	Response.Write "</p></form>"



	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	strSQL="WHERE ([log_Type]=1) AND ([log_Level]>0) AND (1=1) "

	If CheckRights("Root")=False And CheckRights("ArticleAll")=False Then strSQL= strSQL & "AND [log_AuthorID] = " & BlogUser.ID

	If intCate<>-1 Then
		Dim strSubCateID : strSubCateID=Join(GetSubCateID(intCate,True),",")
		strSQL= strSQL & " AND [log_CateID] IN (" & strSubCateID & ")"
	End If

	If intLevel<>-1 Then
		strSQL= strSQL & " AND [log_Level] = " & intLevel
	End If

	If intTitle<>"-1" Then
		If ZC_MSSQL_ENABLE=False Then
			strSQL = strSQL & "AND ( (InStr(1,LCase([log_Title]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('" & intTitle &"'),0)<>0) )"
		Else
			strSQL = strSQL & "AND ( (CHARINDEX('" & intTitle &"',[log_Title])<>0) OR (CHARINDEX('" & intTitle &"',[log_Intro])<>0) OR (CHARINDEX('" & intTitle &"',[log_Content])<>0) )"
		End If
	End If

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0""  class=""tableBorder tableBorder-thcenter"">"
	Response.Write "<tr><th width='5%'>"& ZC_MSG076 &"</th><th width='14%'>"& ZC_MSG003 &"</th><th>"& ZC_MSG060 &"</th><th width='14%'>"& ZC_MSG075 &"</th><th width=""6%"">"& ZC_MSG013 &"</th><th width=""9%"">"& ZC_MSG061 &"</th><th width=""12%""></th></tr>"

	objRS.Open("SELECT * FROM [blog_Article] "& strSQL &" ORDER BY [log_PostTime] DESC")
	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			Response.Write "<tr>"

			Response.Write "<td>" & objRS("log_ID") & "</td>"

			Call GetUsersbyUserIDList(objRS("log_AuthorID"))
			Dim User
			For Each User in Users
				If IsObject(User) Then
					If User.ID=objRS("log_AuthorID") Then
						Response.Write "<td>" & User.Name & "</td>"
					End If
				End If
			Next

			Response.Write "<td><div style='overflow:hidden;height:1.5em;'><a href=""../../view.asp?nav=" & objRS("log_ID") & """ title="""& Replace(objRS("log_Title"),"""","") &""" target=""_blank"">" & objRS("log_Title") & "</a></div></td>"
			Response.Write "<td>" & FormatDateTime(objRS("log_PostTime"),vbShortDate) & "</td>"
			Response.Write "<td>" & objRS("log_CommNums") & "</td>"
			Response.Write "<td>" & ZVA_Article_Level_Name(objRS("log_Level")) & "</td>"
			Response.Write "<td align=""center""><a href=""../cmd.asp?act=ArticleEdt&amp;type=Page&amp;webedit="& ZC_BLOG_WEBEDIT &"&amp;id=" & objRS("log_ID") & """><img src=""../image/admin/page_edit.png"" alt=""" & ZC_MSG100 & """ title=""" & ZC_MSG100 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			Response.Write "<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=ArticleDel&amp;type=Page&amp;id=" & objRS("log_ID") & """><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	Response.Write "</table>"

	If  intPageAll>1 Then 
		strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"../admin/admin.asp?act=ArticleMng&amp;type=Page&amp;cate="&ReQuest("cate")&"&amp;level="&ReQuest("level")&"&amp;title="&Escape(ReQuest("title")) & "&amp;page=")
		Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	End If 

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

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class='tableBorder tableBorder-thcenter'>"
	Response.Write "<tr><th width=""5%""></th><th width=""10%"">"& ZC_MSG076 &"</th><th width=""10%"">"& ZC_MSG079 &"</th><th>"& ZC_MSG001 &"</th><th>"& ZC_MSG147 &"</th><th width=""14%""></th></tr>"


	Response.Write "<tr><td align=""center""><img width=""16"" src=""../image/admin/folder.png"" alt="""" /></td>"
	Response.Write "<td>" & Categorys(0).ID & "</td>"
	Response.Write "<td>" & Categorys(0).Order & "</td>"
	Response.Write "<td><a href="""& Categorys(0).Url &"""  target=""_blank"">" & Categorys(0).Name & "</a></td>"
	Response.Write "<td>" & Categorys(0).Alias & "</td>"
	Response.Write "<td align=""center""><a href=""../cmd.asp?act=CategoryEdt&amp;id="& Categorys(0).ID &"""><img src=""../image/admin/folder_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a></td>"
	Response.Write "</tr>"




	Dim aryCateInOrder
	aryCateInOrder=GetCategoryOrder()

	If IsArray(aryCateInOrder) Then
	For i=LBound(aryCateInOrder)+1 To Ubound(aryCateInOrder)

		If Categorys(aryCateInOrder(i)).ParentID=0 Then

			Response.Write "<tr><td align=""center""><img width=""16"" src=""../image/admin/folder.png"" alt="""" /></td>"
			Response.Write "<td>" & Categorys(aryCateInOrder(i)).ID & "</td>"
			Response.Write "<td>" & Categorys(aryCateInOrder(i)).Order & "</td>"
			Response.Write "<td><a href="""& Categorys(aryCateInOrder(i)).Url &"""  target=""_blank"">" & Categorys(aryCateInOrder(i)).Name & "</a></td>"
			Response.Write "<td>" & Categorys(aryCateInOrder(i)).Alias & "</td>"
			Response.Write "<td align=""center""><a href=""../cmd.asp?act=CategoryEdt&amp;id="& Categorys(aryCateInOrder(i)).ID &"""><img src=""../image/admin/folder_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=CategoryDel&amp;id="& Categorys(aryCateInOrder(i)).ID &"""><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
			Response.Write "</tr>"

			For j=1 To UBound(aryCateInOrder)

				If Categorys(aryCateInOrder(j)).ParentID=Categorys(aryCateInOrder(i)).ID Then
					Response.Write "<tr><td align=""center""><img width=""16"" src=""../image/admin/arrow_turn_right.png"" alt="""" /></td>"
					Response.Write "<td>" & Categorys(aryCateInOrder(j)).ID & "</td>"
					Response.Write "<td>" & Categorys(aryCateInOrder(j)).Order & "</td>"
					Response.Write "<td><a href="""& Categorys(aryCateInOrder(j)).Url &"""  target=""_blank"">" & Categorys(aryCateInOrder(j)).Name & "</a></td>"
					Response.Write "<td>" & Categorys(aryCateInOrder(j)).Alias & "</td>"
					Response.Write "<td align=""center""><a href=""../cmd.asp?act=CategoryEdt&amp;id="& Categorys(aryCateInOrder(j)).ID &"""><img src=""../image/admin/folder_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=CategoryDel&amp;id="& Categorys(aryCateInOrder(j)).ID &"""><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
					Response.Write "</tr>"
				End If

			Next

		End If

	Next
	End If

	Response.Write "</table>"

	Response.Write "<p>&nbsp;</p>"
	Response.Write "</div>"

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aCategoryMng"");</script>"

	ExportCategoryList=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Manager Comments
'*********************************************************
Function ExportCommentList(intPage,intContent,isCheck)



	Dim ArtDic

	Set ArtDic=CreateObject("Scripting.Dictionary")

	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim intPageAll

	Call CheckParameter(intPage,"int",1)
	Call CheckParameter(isCheck,"bool",False)
	intContent=FilterSQL(intContent)

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""
	
	Call Add_Response_Plugin("Response_Plugin_CommentMng_SubMenu",MakeSubMenu(ZC_MSG097,"admin.asp?act=CommentMng&amp;page=","m-left" & IIf(isCheck,""," m-now"),False))
	Dim objRS1
	Set objRS1=objConn.Execute("SELECT COUNT([comm_ID]) FROM [blog_Comment] WHERE [comm_isCheck]=-1 Or [comm_isCheck]=1" & IIf(CheckRights("Root"),""," And [comm_AuthorID]=" & BlogUser.ID))

	Dim strtmpresponse
	strtmpresponse=ZC_MSG104
	If (Not objRS1.bof) And (Not objRS1.eof) Then
		strtmpresponse=strtmpresponse&" ("&objRS1(0)&")"
	End If
	Set objRs1=Nothing
	
	Call Add_Response_Plugin("Response_Plugin_CommentMng_SubMenu",MakeSubMenu(strtmpresponse,"admin.asp?act=CommentMng&amp;isCheck=True","m-left" & IIf(isCheck," m-now",""),False))
	
	If isCheck Then
		strSQL=strSQL&" WHERE  ([log_ID]>0) AND ([comm_isCheck]<>0) "
	Else
		strSQL=strSQL&" WHERE  ([log_ID]>0) AND ([comm_isCheck]=0) "
	End If
	
	If CheckRights("Root")=False And CheckRights("CommentAll")=False Then
		strSQL=strSQL & "AND( ([comm_AuthorID] = " & BlogUser.ID & " ) OR ((SELECT [log_AuthorID] FROM [blog_Article] WHERE [blog_Article].[log_ID]=[blog_Comment].[log_ID])=" & BlogUser.ID & " )) "
	End If

	If Trim(intContent)<>"" Then
		strSQL=strSQL & " AND ( ([comm_Author] LIKE '%" & intContent & "%') OR ([comm_Content] LIKE '%" & intContent & "%') OR ([comm_HomePage] LIKE '%" & intContent & "%') ) "
	End If

	Response.Write "<div class=""divHeader"">" & ZC_MSG068 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_CommentMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"


	Response.Write "<form class=""search"" id=""edit"" method=""post"" action=""../admin/admin.asp?act=CommentMng&amp;isCheck="&isCheck&""">"
	Response.Write "<p>"&ZC_MSG234&":&nbsp;&nbsp;&nbsp;&nbsp;"

	Response.Write "<input id=""intContent"" name=""intContent"" style=""width:250px;"" type=""text"" value="""" /> "
	Response.Write "<input type=""submit"" class=""button"" value="""&ZC_MSG087&"""/>"

	Response.Write "</p></form>"

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder tableBorder-thcenter"">"
	Response.Write "<tr><th width=""5%"">"& ZC_MSG076 &"</th><th width=""6%"">"&ZC_MSG152&"</th><th width='10%'>"& ZC_MSG003 &"</th><th>"& ZC_MSG055 &"</th><th width=""15%"">"& ZC_MSG048 &"</th><th width='18%'></th><th width='5%'  align='center'><a href='' onclick='BatchSelectAll();return false'>"& ZC_MSG229 &"</a></th></tr>"'

	objRS.Open("SELECT * FROM [blog_Comment] "& strSQL &" ORDER BY [comm_ID] DESC")


	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize
			Dim objArticle
			Set objArticle=New TArticle
			If ArtDic.Exists(CLng(objRs("log_ID")))=False Then
				objArticle.LoadInfoById objRs("log_ID")
				ArtDic.Add CLng(objRs("log_ID")), objArticle
			Else
				Set objArticle=ArtDic.Item(CLng(objRs("log_ID")))
			End If

			Response.Write "<tr>"
			Response.Write "<td>" & objRS("comm_ID") & "</td>"
			Response.Write "<td>"&IIF(objRs("comm_ParentID")>0,objRs("comm_ParentID"),"")&"</td>"
			If Trim(objRS("comm_Email"))="" Then
			Response.Write "<td>"& objRS("comm_Author") & "</td>"
			Else
			Response.Write "<td><a href=""mailto:"& objRS("comm_Email") &""">" & objRS("comm_Author") & "</a></td>"
			End If

			Response.Write "<td><a href="""&objArticle.URL&"#cmt"&objRS("comm_ID")&""" target=""_blank""><img src=""../image/admin/comment.png"" alt=""" & ZC_MSG212& " @ " & objArticle.HtmlTitle & """ title=""" & ZC_MSG212& " @ " & objArticle.HtmlTitle & """ width=""16"" /></a><a id=""mylink"&objRS("comm_ID")&""" href=""$div"&objRS("comm_ID")&"tip?width=400"" class=""betterTip"" title="""&ZC_MSG055&""">" & Left(objRS("comm_Content"),30) & "</a><div id=""div"&objRS("comm_ID")&"tip"" style=""display:none;""><p>"& objRS("comm_Content") &"</p><br/><p>" & ZC_MSG080 & " : " &objRS("comm_IP") & "</p><p>" & ZC_MSG075 & " : " &objRS("comm_PostTime") & "</p></div></td>"
			Response.Write "<td><div style='overflow:hidden;height:1.5em;'>"& Left(objArticle.HtmlTitle,18) &"</div></td>"
			Response.Write "<td align=""center""><a href=""../cmd.asp?act=CommentEdt&amp;id=" & objRS("comm_ID") &"&amp;revid="&objRs("comm_ID")&"&amp;log_id="& objRS("log_ID") &"""><img src=""../image/admin/comments.png"" alt=""" & ZC_MSG149 & """ title=""" & ZC_MSG149 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=""../cmd.asp?act=CommentEdt&amp;id=" & objRS("comm_ID") & "&amp;log_id="& objRS("log_ID") &"&amp;revid=0""><img src=""../image/admin/comment_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=""../cmd.asp?act=CommentDel&amp;id=" & objRS("comm_ID")  &""" onclick='return window.confirm("""& ZC_MSG058 &""");'><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a>"
			
			Response.Write IIf(CheckRights("Root"),"&nbsp;&nbsp;&nbsp;&nbsp;<a href=""../cmd.asp?act=CommentAudit&amp;id="&objRs("comm_ID")&"""><img src=""../image/admin/"&IIf(isCheck,"ok.png","minus-shield.png")&""" alt="""&IIf(isCheck,ZC_MSG091,ZC_MSG092)&""" title="""&IIf(isCheck,ZC_MSG091,ZC_MSG092)&""" width=""16""/></a>","")
			Response.Write "</td>"
			Response.Write "<td align=""center"" ><input type=""checkbox"" id=""edtDel"&objRS("comm_ID")&""" name=""edtDel"" value="""&objRS("comm_ID")&"""/></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

			Set objArticle=NoThing

		Next
	Set objArticle=Nothing
	End If

	Response.Write "</table>"


	Response.Write "<form id=""frmBatch"" style=""float:left;"" method=""post"" action=""../cmd.asp?act=CommentDelBatch""><input type=""hidden"" id=""edtBatch"" name=""edtBatch"" value=""""/><input class=""button"" type=""submit"" onclick='BatchDeleteAll(""edtBatch"");if(document.getElementById(""edtBatch"").value){return window.confirm("""& ZC_MSG058 &""");}else{return false}' value="""&ZC_MSG228&""" id=""btnPost""/>&nbsp;&nbsp;&nbsp;&nbsp;</form>" & vbCrlf
	
	Response.Write IIf(CheckRights("Root"),"<form id=""frmBatch2"" style=""float:left;"" method=""post"" action=""../cmd.asp?act=CommentAudit""><input type=""hidden"" id=""edtBatch2"" name=""edtBatch"" value=""""/><input class=""button"" type=""submit"" onclick='BatchDeleteAll(""edtBatch2"");if(document.getElementById(""edtBatch2"").value){return window.confirm("""& ZC_MSG058 &""");}else{return false}' value="""&IIf(isCheck,ZC_MSG174,ZC_MSG177)&""" id=""btnPost2""/>&nbsp;&nbsp;&nbsp;&nbsp;"&IIf(isCheck,"<input class=""button"" type=""submit"" onclick='if(window.confirm("""& ZC_MSG058 &""")){document.getElementById(""edtBatch2"").value=""delall""}else{return false}' value="""&ZC_MSG222&""" id=""btnPost3""/>","") &"</form>","") &vbCrlf

	Response.Write "<div class=""clear""></div>"

	If  intPageAll>1 Then 
		strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"admin.asp?act=CommentMng&amp;isCheck="&isCheck&"&amp;page=")
		Response.Write "<p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	End If 

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

		Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder tableBorder-thcenter"">"
		Response.Write "<tr><th width='5%'>"& ZC_MSG076 &"</th><th width='10%'></th><th>"& ZC_MSG003 &"</th><th>"& ZC_MSG147 &"</th><th width='10%'>"& ZC_MSG082 &"</th><th width='10%'>"& ZC_MSG124 &"</th><th width='14%'></th></tr>"

		For i=1 to objRS.PageSize

			Response.Write "<tr>"
			Response.Write "<td>" & objRS("mem_ID") & "</td>"
			Response.Write "<td>" & ZVA_User_Level_Name(objRS("mem_Level")) & "</td>"
			Response.Write "<td>" & objRS("mem_Name") & "</td>"
			Response.Write "<td>" & objRS("mem_Url") & "</td>"
			Response.Write "<td>" & objRS("mem_PostLogs") & "</td>"
			Response.Write "<td>" & objRS("mem_PostComms") & "</td>"

			Response.Write "<td align=""center""><a href=""../cmd.asp?act=UserEdt&amp;id="& objRS("mem_ID") &"""><img src=""../image/admin/user_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=UserDel&amp;id="& objRS("mem_ID") &"""><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"

			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

		Response.Write "</table>"

	End If
	
	Response.Write "<p>"& ZC_MSG189 &"</p>"

	If  intPageAll>1 Then 
		strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"admin.asp?act=UserMng&amp;page=")
		Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	End If 

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
	Response.Write "<p><input type=""file"" id=""edtFileLoad"" name=""edtFileLoad"" size=""40"" />&nbsp;&nbsp;&nbsp;&nbsp;<input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" name=""B1"" onclick='document.getElementById(""edit"").action=document.getElementById(""edit"").action+""&amp;filename=""+escape(edtFileLoad.value)' />&nbsp;&nbsp;<input class=""button"" type=""reset"" value="""& ZC_MSG088 &""" name=""B2"" />"
	Response.Write "&nbsp;<input type=""checkbox"" onclick='if(this.checked==true){document.getElementById(""edit"").action=document.getElementById(""edit"").action+""&amp;autoname=1"";}else{document.getElementById(""edit"").action=""../cmd.asp?act=FileUpload"";};SetCookie(""chkAutoFileName"",this.checked,365);' id=""chkAutoName""/><label for=""chkAutoName"">"& ZC_MSG131 &"</label></p></form>"

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	If CheckRights("Root")=False And CheckRights("FileAll")=False Then strSQL="WHERE [ul_AuthorID] = " & BlogUser.ID

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder tableBorder-thcenter"">"
	Response.Write "<tr><th width='5%'>"& ZC_MSG076 &"</th><th width='10%'>"& ZC_MSG003 &"</th><th width=''>"& ZC_MSG001 &"</th><th width='12%'>"& ZC_MSG041 &"</th><th width='12%'>"& ZC_MSG075 &"</th><th width='5%'></th><th width='5%'><a href='' onclick='BatchSelectAll();return false'>"& ZC_MSG229 &"</a></th></tr>"

	objRS.Open("SELECT * FROM [blog_UpLoad] " & strSQL & " ORDER BY [ul_PostTime] DESC")
	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			Response.Write "<tr><td>"&objRS("ul_ID")&"</td>"

			Call GetUsersbyUserIDList(objRS("ul_AuthorID"))
			Dim User
			For Each User in Users
				If IsObject(User) Then
					If User.ID=objRS("ul_AuthorID") Then
						Response.Write "<td>" & User.Name & "</td>"
					End If
				End If
			Next

			Response.Write "<td><a href='"& BlogHost & ZC_UPLOAD_DIRECTORY &"/"&Year(objRS("ul_PostTime")) & "/" & Month(objRS("ul_PostTime")) & "/"&Server.URLEncode(objRS("ul_FileName"))&"' target='_blank'>"&Year(objRS("ul_PostTime")) & "/" & Month(objRS("ul_PostTime")) & "/" &objRS("ul_FileName")&"</a></td>"

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

	If  intPageAll>1 Then
		strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"admin.asp?act=FileMng&amp;page=")
		Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	End If 

	Response.Write "</div><script type=""text/javascript"">if(GetCookie(""chkAutoFileName"")==""true""){document.getElementById(""chkAutoName"").checked=true;document.getElementById(""edit"").action=document.getElementById(""edit"").action+String.fromCharCode(38)+""autoname=1"";};</script>"
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

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder tableBorder-thcenter"">"
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

	If  intPageAll>1 Then
		strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"admin.asp?act=TagMng&amp;page=")
		Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	End If 

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


	Dim f, f1, fc
	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")

	Response.Write "<div class=""divHeader"">" & ZC_MSG107 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_PlugInMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"




	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder tableBorder-thcenter"">"
	Response.Write "<tr><th width=""50px""></th></th><th>"& ZC_MSG001 &"</th><th width=""12%"">"& ZC_MSG128 &"</th><th width=""12%"">"& ZC_MSG151 &"</th><th width=""14%""></th></tr>"

	Dim objXmlFile,strXmlFile



	strXmlFile =BlogPath & "zb_users/theme/" & ZC_BLOG_THEME & "/" & "theme.xml"

	Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
	objXmlFile.async = False
	objXmlFile.ValidateOnParse=False
	objXmlFile.load(strXmlFile)
	If objXmlFile.readyState=4 Then
		If objXmlFile.parseError.errorCode <> 0 Then
		Else

			If CLng(objXmlFile.documentElement.selectSingleNode("plugin/level").text)>0 Then

				If Err.Number=0 Then

					Response.Write "<tr>"
					Response.Write "<td align='center'><img alt='' width='32' src='"&BlogHost & "zb_users/theme/"& ZC_BLOG_THEME &"/ScreenShot.png' style='margin:2px;'/></td>"
					'Response.Write "<td>"& "0" &"</td>"
					Response.Write "<td><a id=""mylink"&Left(md5(objXmlFile.documentElement.selectSingleNode("id").text),6)&""" href=""$div"&Left(md5(objXmlFile.documentElement.selectSingleNode("id").text),6)&"tip?width=300"" class=""betterTip"" title=""$content"">" & "" & objXmlFile.documentElement.selectSingleNode("name").text & " ("& ZC_MSG199 &")&nbsp;&nbsp;&nbsp;" & objXmlFile.documentElement.selectSingleNode("version").text  & "</a><div id=""div"&Left(md5(objXmlFile.documentElement.selectSingleNode("id").text),6)&"tip"" style=""display:none;"">"&objXmlFile.documentElement.selectSingleNode("note").text&"</div></td>"
					Response.Write "<td>" & "<a target=""_blank"" href=""" & objXmlFile.documentElement.selectSingleNode("author/url").text & """>"& objXmlFile.documentElement.selectSingleNode("author/name").text & "</td>"
					'Response.Write "<td>" & objXmlFile.documentElement.selectSingleNode("version").text & "</td>"
					Response.Write "<td>"& objXmlFile.documentElement.selectSingleNode("modified").text &"</td>"
					Response.Write "<td align='center'>"
					If BlogUser.Level<=CLng(objXmlFile.documentElement.selectSingleNode("plugin/level").text) Then
						If PublicObjFSO.FileExists(BlogPath & "zb_users/theme/" & ZC_BLOG_THEME & "/plugin/" & objXmlFile.documentElement.selectSingleNode("plugin/path").text) Then
							Response.Write "<a href=""../../ZB_USERS/theme/" & ZC_BLOG_THEME & "/plugin/" & objXmlFile.documentElement.selectSingleNode("plugin/path").text &"""><img width='16' title='"&ZC_MSG022&"' alt='"&ZC_MSG022&"' src='../IMAGE/ADMIN/setting_tools.png'/></a>"
						End If
					End If
					Response.Write "</td>"
					Response.Write "</tr>"

				End If

			End If

		End If
	End If
	Set objXmlFile=Nothing

	Set f = PublicObjFSO.GetFolder(BlogPath & "zb_users/plugin/")
	Set fc = f.SubFolders
	For Each f1 in fc

		s=""

		If PublicObjFSO.FileExists(BlogPath & "zb_users/plugin/" & f1.name & "/" & "plugin.xml") Then

			strXmlFile =BlogPath & "zb_users/plugin/" & f1.name & "/" & "plugin.xml"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else
					'If BlogUser.Level<=CLng(objXmlFile.documentElement.selectSingleNode("level").text) Then

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
				s=s & "<td align='center' class='plugin plugin-on'>"
			Else
				s=s & "<td align='center' class='plugin'>"
			End If

			If CheckPluginState(objXmlFile.documentElement.selectSingleNode("id").text) Then

				If PublicObjFSO.FileExists(BlogPath & "zb_users/plugin/" & f1.name & "/" & "logo.png") Then
					s=s & "<img alt='' width='32' src='"&BlogHost & "zb_users/plugin/" & f1.name & "/" & "logo.png"&"'/ style='margin:2px;'>"
				Else
					s=s & "<img alt='' width='32' src='../IMAGE/ADMIN/app-logo.png'/ style='margin:2px;'>"
				End If
			Else
				If PublicObjFSO.FileExists(BlogPath & "zb_users/plugin/" & f1.name & "/" & "logo.png") Then
					s=s & "<img style=""opacity:0.2"" alt='' width='32' src='"&BlogHost & "zb_users/plugin/" & f1.name & "/" & "logo.png"&"'/ style='margin:2px;'>"
				Else
					s=s & "<img style=""opacity:0.2"" alt='' width='32' src='../IMAGE/ADMIN/app-logo.png'/ style='margin:2px;'>"
				End If
			End If

			s=s & "<strong style='display:none;'>"& Server.URLEncode(objXmlFile.documentElement.selectSingleNode("id").text) &"</strong>"

			s=s & "</td>"

			's=s & "<td>"& m &"</td>"
			s=s & "<td><a id=""mylink"&Left(md5(objXmlFile.documentElement.selectSingleNode("id").text),6)&""" href=""$div"&objXmlFile.documentElement.selectSingleNode("id").text&"tip?width=300"" class=""betterTip"" title=""$content"">" & "" & objXmlFile.documentElement.selectSingleNode("name").text & "&nbsp;&nbsp;&nbsp;" & objXmlFile.documentElement.selectSingleNode("version").text & "</a><div id=""div"&objXmlFile.documentElement.selectSingleNode("id").text&"tip"" style=""display:none;"">"&objXmlFile.documentElement.selectSingleNode("note").text&"</div></td>"
			s=s & "<td>" & "<a target=""_blank"" href=""" & objXmlFile.documentElement.selectSingleNode("author/url").text & """>"& objXmlFile.documentElement.selectSingleNode("author/name").text & "</a></td>"
			's=s & "<td>" & objXmlFile.documentElement.selectSingleNode("version").text & "</td>"
			s=s & "<td>"& objXmlFile.documentElement.selectSingleNode("modified").text &"</td>"

				s=s & "<td align='center'>"
			If CheckPluginState(objXmlFile.documentElement.selectSingleNode("id").text) Then
				If CheckRights("PlugInDisable")=True Then
					s=s & "<a href=""../cmd.asp?act=PlugInDisable&amp;name="& Server.URLEncode(objXmlFile.documentElement.selectSingleNode("id").text) &"""><img width='16' title='"&ZC_MSG203&"' alt='"&ZC_MSG203&"' src='../IMAGE/ADMIN/control-power.png'/></a>"
				Else

				End If
			Else
				If CheckRights("PlugInActive")=True Then
					s=s & "<a href=""../cmd.asp?act=PlugInActive&amp;name="& Server.URLEncode(objXmlFile.documentElement.selectSingleNode("id").text) &"""><img width='16' title='"&ZC_MSG202&"' alt='"&ZC_MSG202&"' src='../IMAGE/ADMIN/control-power-off.png'/></a>"
				Else
				End If
			End If

			If CheckPluginState(objXmlFile.documentElement.selectSingleNode("id").text) Then
				If BlogUser.Level<=CLng(objXmlFile.documentElement.selectSingleNode("level").text) Then
					If PublicObjFSO.FileExists(BlogPath & "zb_users/plugin/" & f1.name & "/" & objXmlFile.documentElement.selectSingleNode("path").text) Then
						s=s & "&nbsp;&nbsp;&nbsp;&nbsp;<a href=""../../ZB_USERS/plugin/" & f1.name & "/" & objXmlFile.documentElement.selectSingleNode("path").text &"""><img width='16' title='"&ZC_MSG022&"' alt='"&ZC_MSG022&"' src='../IMAGE/ADMIN/setting_tools.png'/></a>"
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

	Dim s,k


	Response.Write "<div class=""divHeader"">" & ZC_MSG159 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_SiteInfo_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"
	
	
	If BlogUser.Level<4 Then 
		s=s & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"" width=""100%"" class=""tableBorder"" id=""tbStatistic""><tr><th height=""32"" colspan=""4""  align=""center"">&nbsp;"&ZC_MSG167&"&nbsp;<a href=""javascript:statistic('?reload');"">["&ZC_MSG225&ZC_MSG281&"]</a> <img id=""statloading"" style=""display:none"" src=""../image/admin/loading.gif""></th></tr><tr><td></td></tr></table>"
	End If
	If Len(ZC_UPDATE_INFO_URL)>0 Then
		s=s & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"" width=""100%"" class=""tableBorder""><tr><th height=""32"" colspan=""4"" align=""center"">&nbsp;"&ZC_MSG164&"&nbsp;<a href=""javascript:updateinfo('?reload');"">["&ZC_MSG225&"]</a> <img id=""infoloading"" style=""display:none"" src=""../image/admin/loading.gif""></th></tr><tr><td height=""25"" colspan=""4"" id=""tdUpdateInfo"">&nbsp;</td></tr></table>"
	End If
	k = LoadFromFile(BlogPath & "zb_system\defend\thanks.html","utf-8")
	k = Replace(k,"{%ZC_MSG303%}",ZC_MSG303)
	k = Replace(k,"{%ZC_MSG304%}",ZC_MSG304)
	k = Replace(k,"{%ZC_MSG305%}",ZC_MSG305)
	k = Replace(k,"{%ZC_MSG306%}",ZC_MSG306)
	k = Replace(k,"{%ZC_MSG307%}",ZC_MSG307)
	k = Replace(k,"{%ZC_MSG308%}",ZC_MSG308)
	k = Replace(k,"{%ZC_MSG309%}",ZC_MSG309)
	s = s & k
	Response.Write s
	Response.Write Response_Plugin_Admin_SiteInfo
	Response.Write "</div>"

	Response.Write "<script type=""text/javascript"">statistic("""");updateinfo("""");</script>"
	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aSiteInfo"");</script>"
	Response.Write "<script type=""text/javascript"">ActiveTopMenu(""topmenu1"");</script>"

	ExportSiteInfo=True

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function ExportFileReBuildAsk()

	Response.Write "<div class=""divHeader2"">" & ZC_MSG073 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_AskFileReBuild_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"

	Response.Write "<form id=""edit"" name=""edit"" method=""post"" action=""../cmd.asp?act=FileReBuild"">" & vbCrlf
	Response.Write "<p>"& ZC_MSG112 &"</p>" & vbCrlf

	Response.Write "<p><input class=""button"" type=""submit"" value="""&ZC_MSG087&""" id=""btnPost""/>"

	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input style=""display:none;"" class=""button"" type=""button"" onclick='$(this).prop({disabled: true});BatchCancel()' value="""&ZC_MSG264&"""/>"

	Response.Write "<script type=""text/javascript"">if(window.webkitNotifications&&(!window.webkitNotifications.checkPermission() == 0)){document.write('&nbsp;&nbsp;&nbsp;&nbsp;<input class=""button"" onclick=""window.webkitNotifications.requestPermission();return false;"" type=""button"" value="""&ZC_MSG263&"""/>')}</script>"	

	Response.Write "</p>" & vbCrlf
	Response.Write "</form>" 

	Response.Write "</div>"



	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aAskFileReBuild"");</script>"
	Response.Write "<script type=""text/javascript"">ActiveTopMenu(""topmenu3"");</script>"

	Response.Write "<script type=""text/javascript"">function BatchBegin(){$(""input[type='submit']"").prop({disabled: true});$(""input[type='button']"").show();};</script>"
	Response.Write "<script type=""text/javascript"">function BatchEnd(){$(""input[type='submit']"").prop({disabled: false});$(""input[type='button']"").hide();};</script>"

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
	Dim Theme_Modified
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


	Response.Write "<div class=""divHeader"">" & ZC_MSG223 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_ThemeMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"



	Response.Write "<form id=""frmTheme"" method=""post"" action=""../cmd.asp?act=ThemeSav"">"

	Dim objXmlFile,strXmlFile
	Dim f, f1, fc, s
	
	If Not IsObject(PublicObjFSO) Then Set PublicObjFSO=Server.CreateObject("Scripting.FileSystemObject")
	
	Set f = PublicObjFSO.GetFolder(BlogPath & "zb_users/theme" & "/")
	Set fc = f.SubFolders
	For Each f1 in fc

		If PublicObjFSO.FileExists(BlogPath & "zb_users/theme" & "/" & f1.name & "/" & "theme.xml") Then

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
					Theme_Modified=""
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
					Theme_Modified=objXmlFile.documentElement.selectSingleNode("modified").text

					Theme_Description=objXmlFile.documentElement.selectSingleNode("description").text

					Theme_ScreenShot="../../zb_users/theme" &"/" & Theme_Id & "/" & "screenshot.png"




		If UCase(Theme_Id)=UCase(CurrentTheme) Then
			Response.Write "<div id=""theme-"&Theme_Id&""" class=""theme theme-now"">"
		Else
			Response.Write "<div id=""theme-"&Theme_Id&""" class=""theme theme-other"">"
		End If

		If UCase(Theme_Id) <> UCase(f1.name) Then
			Response.Write "<div style=""color:red;"">ID Error! Should be ""<strong>"& f1.name &"</strong>""!!</div>"
		Else
			Response.Write "<div class=""theme-name""><img width='16' title='' alt='' src='../IMAGE/ADMIN/layout.png'/> <a  target=""_blank"" href="""&Theme_Url&"""  title="""">" & "<strong style='display:none;'>" & Server.URLEncode(Theme_Id) & "</strong><b>" & Theme_Name & "</b>" & "</a>"
			If UCase(Theme_Id)=UCase(CurrentTheme) Then
				If PublicObjFSO.FileExists(BlogPath & "zb_users/theme/" & ZC_BLOG_THEME & "/plugin/" & objXmlFile.documentElement.selectSingleNode("plugin/path").text) Then
					Response.Write "<input type=""button"" class=""theme-config button"" value="""&ZC_MSG278&""" onclick=""location.href='"&BlogHost&"zb_users/theme/"&ZC_BLOG_THEME&"/plugin/"&objXmlFile.documentElement.selectSingleNode("plugin/path").text&"'"">"
				End If
			End If
			Response.Write "</div>"
		End If


		Response.Write "<div><a id=""mylink"&Left(md5(Theme_Id),6)&""" href=""$div"&Left(md5(Theme_Id),6)&"tip?width=320"" class=""betterTip"" title="""&Theme_Name&""" "
		If UCase(Theme_Id)<>UCase(CurrentTheme) Then Response.Write " onclick='$(""#edtZC_BLOG_THEME"").val("""&Theme_Id&""");$(""#edtZC_BLOG_CSS"").val($(""#cate"&Left(md5(Theme_Id),6)&""").val());$(""#frmTheme"").submit();'"
		Response.Write "><img src=""" & Theme_ScreenShot & """ alt=""ScreenShot"" width=""200"" height=""150"" /></a></div>"

		Response.Write "<div id=""div"&Left(md5(Theme_Id),6)&"tip"" style=""display:none;"">"
		Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"" width=""100%"" class=""tableBorder""><tbody>"
		Response.Write "<tr><th colspan=""2"">ID : " & Theme_Id & "</th></tr>"
		Response.Write "<tr><td width=""60px"">"&ZC_MSG128&"</td><td>" & Theme_Author_Name &"</td></tr>"
		Response.Write "<tr><td>"&ZC_MSG054&"</td><td>" & Theme_Author_Url & "</td></tr>"
		Response.Write "<tr><td>"&ZC_MSG197&"</td><td>" & Theme_Source_Name & "</td></tr>"
		Response.Write "<tr><td>"&ZC_MSG054&"</td><td>" & Theme_Source_Url & "</td></tr>"
		Response.Write "<tr><td>"&ZC_MSG011&"</td><td>" & Theme_Pubdate & "</td></tr>"
		Response.Write "<tr><td>"&ZC_MSG151&"</td><td>" & Theme_Modified & "</td></tr>"
		Response.Write "<tr><td>"&ZC_MSG198&"</td><td>" & TransferHTML(Theme_Description,"[enter]") & "</tr>"
		Response.Write "</tbody></table>"
		Response.Write "</div>"

'		If Theme_Url="" Then
'			Response.Write "<p>"&ZC_MSG001&":" & Theme_Name & "</p>"
'		Else
'			Response.Write "<p>"&ZC_MSG001&":<a target=""_blank"" href=""" & Theme_Url & """>" & Theme_Name & "</a></p>"
'		End If

		If Theme_Author_Url="" Then
			Response.Write "<div class=""theme-author"">"&ZC_MSG128&": " & Theme_Author_Name & "</div>"
		Else
			Response.Write "<div class=""theme-author"">"&ZC_MSG128&": <a target=""_blank"" href=""" & Theme_Author_Url & """>" & Theme_Author_Name & "</a></div>"
		End If


'		Response.Write "<p>"&ZC_MSG011&":" & Theme_Pubdate & "</p>"
'		Response.Write "<p style='height:1.0em;'>"&ZC_MSG016&":" & Theme_Note & "</p>"
		Response.Write "<div class=""theme-style"">"&ZC_MSG196&": " & "<select class=""edit"" size=""1"" id=""cate"&Left(md5(Theme_Id),6)&""" name=""cate"&Left(md5(Theme_Id),6)&""" style=""width:110px;"" onchange=""document.getElementById('edtZC_BLOG_THEME').value='"&Theme_Id&"';document.getElementById('edtZC_BLOG_CSS').value=this.options[this.selectedIndex].value"">"


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
							Response.Write " <option value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
						ElseIf LCase(Theme_Style_Name)="style" Then
							Response.Write " <option value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
						ElseIf LCase(Theme_Style_Name)=LCase(Theme_Id) Then
							Response.Write " <option value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
						Else
							If i=1 Then
								Response.Write " <option  value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
							Else
								Response.Write " <option value="""& Theme_Style_Name &""">"&aryFileList(i)&"</option> "
							End If
						End If
					End If
				End If
			Next
		End If

		Response.Write "</select>"
		Response.Write "<input type=""button"" class=""theme-activate button"" value="""&ZC_MSG202&""" onclick='if(!document.getElementById(""cate"&Left(md5(Theme_Id),6)&""").value){return false;}else{$(""#edtZC_BLOG_THEME"").val("""&Theme_Id&""");$(""#edtZC_BLOG_CSS"").val($(""#cate"&Left(md5(Theme_Id),6)&""").val());};$(""#frmTheme"").submit()'></div>"


		Response.Write "</div>"



				End If
			Set objXmlFile=Nothing
			End If

		End If

	Next

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

	Call GetFunction()

	Call Add_Response_Plugin("Response_Plugin_FunctionMng_SubMenu",MakeSubMenu(ZC_MSG142 & "","../cmd.asp?act=FunctionEdt","m-left",False))

Call Add_Response_Plugin("Response_Plugin_FunctionMng_SubMenu",MakeSubMenu(ZC_MSG052 & "","../cmd.asp?act=FunctionEdt&amp;id="&Functions(FunctionMetas.GetValue("navbar")).ID,"m-left",False))
Call Add_Response_Plugin("Response_Plugin_FunctionMng_SubMenu",MakeSubMenu(ZC_MSG030 & "","../cmd.asp?act=FunctionEdt&amp;id="&Functions(FunctionMetas.GetValue("favorite")).ID,"m-left",False))
Call Add_Response_Plugin("Response_Plugin_FunctionMng_SubMenu",MakeSubMenu(ZC_MSG031 & "","../cmd.asp?act=FunctionEdt&amp;id="&Functions(FunctionMetas.GetValue("link")).ID,"m-left",False))
Call Add_Response_Plugin("Response_Plugin_FunctionMng_SubMenu",MakeSubMenu(ZC_MSG039 & "","../cmd.asp?act=FunctionEdt&amp;id="&Functions(FunctionMetas.GetValue("misc")).ID,"m-left",False))


	Dim i,j,s,t
	Dim a,b,c,d,e,f
	Response.Write "<div class=""divHeader"">" & ZC_MSG007 & "</div>"
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_FunctionMng_SubMenu & "</div>"
	Response.Write "<div id=""divMain2"">"
	'widget-list begin
	Response.Write "<div class=""widget-left"">"
	Response.Write "<div class=""widget-list"">"

	Response.Write "<script type=""text/javascript"">"
	Response.Write "var functions = {"
		For i=LBound(Functions)+1 To Ubound(Functions)
			If IsObject(Functions(i)) Then
				Response.Write "'"&Functions(i).FileName&"' : '"&Functions(i).SourceType&"'"
				If i<>Ubound(Functions) Then Response.Write ","
			End If
		Next
	Response.Write "};"
	Response.Write "</script>"

	Response.Write "<div class=""widget-list-header"">" & ZC_MSG277 & "</div>"
	Response.Write "<div class=""widget-list-note"">"&ZC_MSG145&"</div>" & vbCrlf
	For i=LBound(Functions)+1 To Ubound(Functions)
		If IsObject(Functions(i)) Then
		If Functions(i).IsSystem Then
			Response.Write "<div class=""widget widget_ishidden_"&LCase(Functions(i).IsHidden)&" widget_source_"& Functions(i).SourceType &" widget_id_" & Functions(i).FileName & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& TransferHTML(Functions(i).Name,"[html-format]")

			Response.Write "	<span class=""widget-action""><a href=""../cmd.asp?act=FunctionEdt&amp;id="&Functions(i).ID&"""><img class=""edit-action"" src=""../image/admin/brick_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>"

			Response.Write "	</span>"
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& Functions(i).FileName &"</div>"	
			Response.Write "</div>"
		End If
		End If
	Next

	Response.Write "<div class=""widget-list-header"">" & ZC_MSG286 & "</div>"
	For i=LBound(Functions)+1 To Ubound(Functions)
		If IsObject(Functions(i)) Then
		If Functions(i).IsUsers Then
			Response.Write "<div class=""widget widget_ishidden_"&LCase(Functions(i).IsHidden)&" widget_source_"& Functions(i).SourceType &" widget_id_" & Functions(i).FileName & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& TransferHTML(Functions(i).Name,"[html-format]")

			Response.Write "	<span class=""widget-action""><a href=""../cmd.asp?act=FunctionEdt&amp;id="&Functions(i).ID&"""><img class=""edit-action"" src=""../image/admin/brick_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>"

			Response.Write "&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=FunctionDel&amp;id="& Functions(i).ID &"""><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a>"

			Response.Write "	</span>"
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& Functions(i).FileName &"</div>"	
			Response.Write "</div>"
		End If
		End If
	Next

	Response.Write "<div class=""widget-list-header"">" & ZC_MSG287 & "</div>"
	For i=LBound(Functions)+1 To Ubound(Functions)
		If IsObject(Functions(i)) Then
		If Functions(i).IsTheme Then
			Response.Write "<div class=""widget widget_ishidden_"&LCase(Functions(i).IsHidden)&" widget_source_"& Functions(i).SourceType &" widget_id_" & Functions(i).FileName & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& TransferHTML(Functions(i).Name,"[html-format]")

			Response.Write "	<span class=""widget-action""><a href=""../cmd.asp?act=FunctionEdt&amp;id="&Functions(i).ID&"""><img class=""edit-action"" src=""../image/admin/brick_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>"

			'If Functions(i).AppName<>ZC_BLOG_THEME Then
				Response.Write "&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=FunctionDel&amp;id="& Functions(i).ID &"""><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a>"
			'End If

			Response.Write "	</span>"
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& Functions(i).FileName &"</div>"	
			Response.Write "</div>"
		End If
		End If
	Next
	Response.Write "<div class=""widget-list-header"">" & ZC_MSG288 & "</div>"
	For i=LBound(Functions)+1 To Ubound(Functions)
		If IsObject(Functions(i)) Then
		If Functions(i).IsPlugin Then
			Response.Write "<div class=""widget widget_ishidden_"&LCase(Functions(i).IsHidden)&" widget_source_"& Functions(i).SourceType &" widget_id_" & Functions(i).FileName & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& TransferHTML(Functions(i).Name,"[html-format]")

			Response.Write "	<span class=""widget-action""><a href=""../cmd.asp?act=FunctionEdt&amp;id="&Functions(i).ID&"""><img class=""edit-action"" src=""../image/admin/brick_edit.png"" alt=""" & ZC_MSG078 & """ title=""" & ZC_MSG078 & """ width=""16"" /></a>"
			
			If Not CheckPluginState(Functions(i).AppName) Then
				Response.Write "&nbsp;<a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""../cmd.asp?act=FunctionDel&amp;id="& Functions(i).ID &"""><img src=""../image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a>"
			End If

			Response.Write "	</span>"
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& Functions(i).FileName &"</div>"	
			Response.Write "</div>"
		End If
		End If
	Next
	Response.Write "<div class=""widget-list-header"">" & ZC_MSG289 & "</div>"
	For i=LBound(Functions)+1 To Ubound(Functions)
		If IsObject(Functions(i)) Then
		If Functions(i).IsOther Then
			Response.Write "<div class=""widget widget_ishidden_"&LCase(Functions(i).IsHidden)&" widget_source_"& Functions(i).SourceType &" widget_id_" & Functions(i).FileName & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& TransferHTML(Functions(i).Name,"[html-format]")

			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& Functions(i).FileName &"</div>"	
			Response.Write "</div>"
		End If
		End If
	Next

	Response.Write "<hr/><form id=""frmBatch"" method=""post"" action=""../cmd.asp?act=FunctionMng"" style=""position: absolute;"">"
	Response.Write "<input type=""hidden"" id=""edtSidebar"" name=""edtSidebar"" value="""&ZC_SIDEBAR_ORDER&"""/>"
	Response.Write "<input type=""hidden"" id=""edtSidebar2"" name=""edtSidebar2"" value="""&ZC_SIDEBAR_ORDER2&"""/>"
	Response.Write "<input type=""hidden"" id=""edtSidebar3"" name=""edtSidebar3"" value="""&ZC_SIDEBAR_ORDER3&"""/>"
	Response.Write "<input type=""hidden"" id=""edtSidebar4"" name=""edtSidebar4"" value="""&ZC_SIDEBAR_ORDER4&"""/>"
	Response.Write "<input type=""hidden"" id=""edtSidebar5"" name=""edtSidebar5"" value="""&ZC_SIDEBAR_ORDER5&"""/>"

	Response.Write "</form>" & vbCrlf
	Response.Write "<div class=""clear""></div></div>"
	Response.Write "</div>"
	'widget-list end

	'siderbar-list begin
	Response.Write "<div class=""siderbar-list"">"
	Response.Write "<div class=""siderbar-drop"" id=""siderbar""><div class=""siderbar-header"">"&ZC_MSG290&"&nbsp;<img class=""roll"" src=""../image/admin/loading.gif"" width=""16"" alt="""" /><span class=""ui-icon ui-icon-triangle-1-s""></span></div><div  class=""siderbar-sort-list"" >"
	t=Split(ZC_SIDEBAR_ORDER,":")	
	Response.Write "<div class=""siderbar-note"" >"&Replace(ZC_MSG295,"%n",UBound(t)+1)&"</div>"
	For Each s In t
		If FunctionMetas.Exists(s)=True Then

			Response.Write "<div class=""widget widget_ishidden_"&LCase(Functions(FunctionMetas.GetValue(s)).IsHidden)&" widget_source_"& Functions(FunctionMetas.GetValue(s)).SourceType &" widget_id_" & Functions(FunctionMetas.GetValue(s)).FileName & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& Functions(FunctionMetas.GetValue(s)).Name 
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& Functions(FunctionMetas.GetValue(s)).FileName &"</div>"	
			Response.Write "</div>"

		Else

			Response.Write "<div class=""widget widget_source_other " & s & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& s
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& s &"</div>"	
			Response.Write "</div>"


		End If
	Next
	Response.Write "</div></div>"

	Response.Write "<div class=""siderbar-drop"" id=""siderbar2""><div class=""siderbar-header"">"&ZC_MSG291&"&nbsp;<img class=""roll"" src=""../image/admin/loading.gif"" width=""16"" alt="""" /><span class=""ui-icon ui-icon-triangle-1-s""></span></div><div  class=""siderbar-sort-list"" >"
	t=Split(ZC_SIDEBAR_ORDER2,":")
	Response.Write "<div class=""siderbar-note"" >"&Replace(ZC_MSG295,"%n",UBound(t)+1)&"</div>"
	For Each s In t
		If FunctionMetas.Exists(s)=True Then

			Response.Write "<div class=""widget widget_ishidden_"&LCase(Functions(FunctionMetas.GetValue(s)).IsHidden)&" widget_source_"& Functions(FunctionMetas.GetValue(s)).SourceType &" widget_id_" & Functions(FunctionMetas.GetValue(s)).FileName & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& Functions(FunctionMetas.GetValue(s)).Name 
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& Functions(FunctionMetas.GetValue(s)).FileName &"</div>"	
			Response.Write "</div>"

		Else

			Response.Write "<div class=""widget widget_source_other " & s & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& s
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& s &"</div>"	
			Response.Write "</div>"


		End If
	Next
	Response.Write "</div></div>"

	Response.Write "<div class=""siderbar-drop"" id=""siderbar3""><div class=""siderbar-header"">"&ZC_MSG292&"&nbsp;<img class=""roll"" src=""../image/admin/loading.gif"" width=""16"" alt="""" /><span class=""ui-icon ui-icon-triangle-1-s""></span></div><div  class=""siderbar-sort-list"" >"
	t=Split(ZC_SIDEBAR_ORDER3,":")
	Response.Write "<div class=""siderbar-note"" >"&Replace(ZC_MSG295,"%n",UBound(t)+1)&"</div>"
	For Each s In t
		If FunctionMetas.Exists(s)=True Then

			Response.Write "<div class=""widget widget_ishidden_"&LCase(Functions(FunctionMetas.GetValue(s)).IsHidden)&" widget_source_"& Functions(FunctionMetas.GetValue(s)).SourceType &" widget_id_" & Functions(FunctionMetas.GetValue(s)).FileName & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& Functions(FunctionMetas.GetValue(s)).Name 
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& Functions(FunctionMetas.GetValue(s)).FileName &"</div>"	
			Response.Write "</div>"

		Else

			Response.Write "<div class=""widget widget_source_other " & s & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& s
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& s &"</div>"	
			Response.Write "</div>"


		End If
	Next
	Response.Write "</div></div>"

	Response.Write "<div class=""siderbar-drop"" id=""siderbar4""><div class=""siderbar-header"">"&ZC_MSG293&"&nbsp;<img class=""roll"" src=""../image/admin/loading.gif"" width=""16"" alt="""" /><span class=""ui-icon ui-icon-triangle-1-s""></span></div><div  class=""siderbar-sort-list"" >"
	t=Split(ZC_SIDEBAR_ORDER4,":")
	Response.Write "<div class=""siderbar-note"" >"&Replace(ZC_MSG295,"%n",UBound(t)+1)&"</div>"
	For Each s In t
		If FunctionMetas.Exists(s)=True Then

			Response.Write "<div class=""widget widget_ishidden_"&LCase(Functions(FunctionMetas.GetValue(s)).IsHidden)&" widget_source_"& Functions(FunctionMetas.GetValue(s)).SourceType &" widget_id_" & Functions(FunctionMetas.GetValue(s)).FileName & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& Functions(FunctionMetas.GetValue(s)).Name 
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& Functions(FunctionMetas.GetValue(s)).FileName &"</div>"	
			Response.Write "</div>"

		Else

			Response.Write "<div class=""widget widget_source_other " & s & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& s
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& s &"</div>"	
			Response.Write "</div>"


		End If
	Next
	Response.Write "</div></div>"

	Response.Write "<div class=""siderbar-drop"" id=""siderbar5""><div class=""siderbar-header"">"&ZC_MSG294&"&nbsp;<img class=""roll"" src=""../image/admin/loading.gif"" width=""16"" alt="""" /><span class=""ui-icon ui-icon-triangle-1-s""></span></div><div  class=""siderbar-sort-list"" >"
	t=Split(ZC_SIDEBAR_ORDER5,":")
	Response.Write "<div class=""siderbar-note"" >"&Replace(ZC_MSG295,"%n",UBound(t)+1)&"</div>"
	For Each s In t
		If FunctionMetas.Exists(s)=True Then

			Response.Write "<div class=""widget widget_ishidden_"&LCase(Functions(FunctionMetas.GetValue(s)).IsHidden)&" widget_source_"& Functions(FunctionMetas.GetValue(s)).SourceType &" widget_id_" & Functions(FunctionMetas.GetValue(s)).FileName & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& Functions(FunctionMetas.GetValue(s)).Name 
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& Functions(FunctionMetas.GetValue(s)).FileName &"</div>"	
			Response.Write "</div>"

		Else

			Response.Write "<div class=""widget widget_source_other " & s & """>"
			Response.Write "	<div class=""widget-title""><img class=""more-action"" width=""16"" src=""../image/admin/brick.png"" alt="""" />"& s
			Response.Write "	</div>"
			Response.Write "	<div class=""funid"" style=""display:none"">"& s &"</div>"	
			Response.Write "</div>"


		End If
	Next
	Response.Write "</div></div>"

	Response.Write "<div class=""clear""></div></div>"
	'siderbar-list end

	Response.Write "<div class=""clear""></div>"

	Response.Write "</div>"

	Response.Write "<script type=""text/javascript"">ActiveLeftMenu(""aFunctionMng"");</script>"

%>
<script type="text/javascript">
	$(function() {

		function sortFunction(){
			var s1="";
			$("#siderbar").find("div.funid").each(function(i){
			   s1 += $(this).html() +":";
			 });

			 var s2="";
			$("#siderbar2").find("div.funid").each(function(i){
			   s2 += $(this).html() +":";
			 });

			 var s3="";
			$("#siderbar3").find("div.funid").each(function(i){
			   s3 += $(this).html() +":";
			 });

			 var s4="";
			$("#siderbar4").find("div.funid").each(function(i){
			   s4 += $(this).html() +":";
			 });

			 var s5="";
			$("#siderbar5").find("div.funid").each(function(i){
			   s5 += $(this).html() +":";
			 });

			$("#edtSidebar" ).val(s1);
			$("#edtSidebar2").val(s2);
			$("#edtSidebar3").val(s3);
			$("#edtSidebar4").val(s4);
			$("#edtSidebar5").val(s5);


			$.post($("#frmBatch").attr("action"),
				{
				"edtSidebar": s1,
				"edtSidebar2": s2,
				"edtSidebar3": s3,
				"edtSidebar4": s4,
				"edtSidebar5": s5
				},
			   function(data){
				 //alert("Data Loaded: " + data);
			   });

		};

		var t;
		function hideWidget(item){
				item.find(".ui-icon").removeClass("ui-icon-triangle-1-s").addClass("ui-icon-triangle-1-w");
				t=item.next();
				t.find(".widget").hide("fast").end().show();
				t.find(".siderbar-note>span").text(t.find(".widget").length);
		}
		function showWidget(item){
				item.find(".ui-icon").removeClass("ui-icon-triangle-1-w").addClass("ui-icon-triangle-1-s");
				t=item.next();
				t.find(".widget").show("fast");
				t.find(".siderbar-note>span").text(t.find(".widget").length);
		}

		$(".siderbar-header").toggle( function () {
				hideWidget($(this));
			  },
			  function () {
				showWidget($(this));
			  });

 		$( ".siderbar-sort-list" ).sortable({
 			items:'.widget',
			start:function(event, ui){
				showWidget(ui.item.parent().prev());
				 var c=ui.item.find(".funid").html();
				 if(ui.item.parent().find(".widget:contains("+c+")").length>1){
					ui.item.remove();
				 };
			} ,			
			stop:function(event, ui){$(this).parent().find(".roll").show("slow");sortFunction();$(this).parent().find(".roll").hide("slow");
				showWidget($(this).parent().prev());
			}
 		}).disableSelection(); 

		$( ".widget-list>.widget" ).draggable({
            connectToSortable: ".siderbar-sort-list",
            revert: "invalid", 
            containment: "document",
            helper: "clone",
            cursor: "move"
        }).disableSelection();

		$( ".widget-list" ).droppable({
			accept:".siderbar-sort-list>.widget",
            drop: function( event, ui ) {
            	ui.draggable.remove();
            }
        });

});

</script>
<%

	ExportFunctionList=True

End Function
'*********************************************************

%>