<%

Function PageMeta_ExportBar(ID)
	dim a,b,c,d,e

	a="<span class=""m-left %1""><a href=""List.asp?act=%2"">[%3]</a> </span>"
	b=Array("","ArticleMng","CategoryMng","UserMng","TagMng")
	c=Array("","文章管理","分类管理","用户管理","Tag管理")
	For e=1 To Ubound(b)
		d=d&a
		if id=e then d=replace(d,"%1","m-now") else d=replace(d,"%1","")
		d=replace(d,"%2",b(e))
		d=replace(d,"%3",c(e))
	Next
	PageMeta_ExportBar=d
End Function
Function PageMeta_ExportArticleList(intPage,intCate,intLevel,intTitle)
	Call GetUser
	Call GetCategory
	Dim i
	Dim objRS
	Dim strSQL,strPage
	Dim intPageAll
	Call CheckParameter(intPage,"int",1)
	Call CheckParameter(intCate,"int",-1)
	Call CheckParameter(intLevel,"int",-1)
	Call CheckParameter(intTitle,"sql",-1)
	intTitle=vbsunescape(intTitle)
	intTitle=FilterSQL(intTitle)
	Call GetBlogHint()
	Response.Write "<form id=""edit"" class=""search"" method=""post"" enctype=""application/x-www-form-urlencoded"" action=""List.asp?act=ArticleMng"">"
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
					Response.Write "<option value="""&Categorys(aryCateInOrder(n)).ID&""">&nbsp;┄ "&TransferHTML( Categorys(aryCateInOrder(n)).Name,"[html-format]")&"</option>"
				End If
			Next
		End If
	Next
	End If
	Response.Write "</select> "
	Response.Write ZC_MSG061&" <select class=""edit"" size=""1"" id=""level"" name=""level"" style=""width:80px;"" ><option value=""-1"">"&ZC_MSG157&"</option> "
	For i=LBound(ZVA_Article_Level_Name)+1 to Ubound(ZVA_Article_Level_Name)
			Response.Write "<option value="""&i&""" "
			Response.Write ">"&ZVA_Article_Level_Name(i)&"</option>"
	Next
	Response.Write "</select>"
	Response.Write " "&ZC_MSG224&" <input id=""title"" name=""title"" style=""width:150px;"" type=""text"" value="""" /> "
	Response.Write "<input type=""submit"" class=""button"" value="""&ZC_MSG087&""">"
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
		If ZC_MSSQL_ENABLE=False Then
			strSQL = strSQL & "AND ( (InStr(1,LCase([log_Title]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('" & intTitle &"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('" & intTitle &"'),0)<>0 ))"
		Else
			strSQL = strSQL & "AND ( (CHARINDEX('" & intTitle &"',[log_Title]))<>0) OR (CHARINDEX('" & intTitle &"',[log_Intro])<>0) OR (CHARINDEX('" & intTitle &"',[log_Content])<>0) "
		End If
	End If
	Response.Write "<table border=""1"" width=""100%"" cellspacing=""1"" cellpadding=""1"">"
	Response.Write "<tr><td>"& ZC_MSG076 &"</td><td>"& ZC_MSG012 &"</td><td>"& ZC_MSG003 &"</td><td>"& ZC_MSG060 &"</td><td>Meta头</td><td></td></tr>"
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
			If Len(objRS("log_Title"))>14 Then
				Response.Write "<td><a href="""&ZC_BLOG_HOST&"/ZB_SYSTEM/view.asp?id=" & objRS("log_ID") & """ title="""& Replace(objRS("log_Title"),"""","") &""" target=""_blank"">" & Left(objRS("log_Title"),14) & "..." & "</a></td>"
			Else
				Response.Write "<td><a href="""&ZC_BLOG_HOST&"/ZB_SYSTEM/view.asp?id=" & objRS("log_ID") & """ title="""& Replace(objRS("log_Title"),"""","") &""" target=""_blank"">" & objRS("log_Title") & "</a></td>"
			End If
			dim mt
			Set mt=New TMeta
			mt.loadstring=objRs("log_Meta")
			
			Response.Write "<td>" & replace(vbsunescape2(mt.getvalue("pagemeta")),vbcrlf,"<br/>") & "</td>"
			set mt=nothing
			Response.Write "<td align=""center""><a href=""Edit.asp?act=1&id=" & objRS("log_ID") & """><img src="""&ZC_BLOG_HOST&"/ZB_SYSTEM/image/admin/folder_edit.png"" alt=""编辑"" title=""编辑"" width=""16"" /></a></td>"
			Response.Write "</tr>"
			objRS.MoveNext
			If objRS.eof Then Exit For
		Next
	End If
	Response.Write "</table> "
	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"List.asp?act=ArticleMng&cate="&ReQuest("cate")&"&amp;level="&ReQuest("level")&"&amp;title="&Escape(ReQuest("title")) & "&amp;page=")

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage

	Response.Write "</p></div>"
	objRS.Close
	Set objRS=Nothing
End Function


Function PageMeta_ExportCategoryList(intPage)

	Dim i,j
	Call GetCategory
	Call CheckParameter(intPage,"int",1)
'∟

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""1"" cellpadding=""1"">"
	Response.Write "<tr><td width=""18""></td><td>"& ZC_MSG076 &"</td><td>"& ZC_MSG079 &"</td><td>"& ZC_MSG001 &"</td><td>"& ZC_MSG147 &"</td><td>Meta头</td><td></td></tr>"

	Dim aryCateInOrder
	aryCateInOrder=GetCategoryOrder()

	If IsArray(aryCateInOrder) Then
	For i=LBound(aryCateInOrder)+1 To Ubound(aryCateInOrder)

		If Categorys(aryCateInOrder(i)).ParentID=0 Then

			Response.Write "<tr><td align=""center""><img width=""16"" src="""&ZC_BLOG_HOST&"ZB_SYSTEM/image/admin/folder.png"" alt="""" /></td>"
			Response.Write "<td>" & Categorys(aryCateInOrder(i)).ID & "</td>"
			Response.Write "<td>" & Categorys(aryCateInOrder(i)).Order & "</td>"
			Response.Write "<td><a href=""../../../catalog.asp?cate="& Categorys(aryCateInOrder(i)).ID &"""  target=""_blank"">" & Categorys(aryCateInOrder(i)).Name & "</a></td>"
			Response.Write "<td>" & Categorys(aryCateInOrder(i)).Alias & "</td>"
			rESPONSE.Write "<td>"&Replace(vbsunescape2(Categorys(aryCateInOrder(i)).Meta.GetValue("pagemeta")),vbcrlf,"<br/>")&"</td>"
			Response.Write "<td align=""center""><a href=""Edit.asp?act=2&id="& Categorys(aryCateInOrder(i)).ID &"""><img src="""&ZC_BLOG_HOST&"/ZB_SYSTEM/image/admin/folder_edit.png"" alt=""编辑"" title=""编辑"" width=""16"" /></a></td>"
			Response.Write "</tr>"

			For j=0 To UBound(aryCateInOrder)

				If Categorys(aryCateInOrder(j)).ParentID=Categorys(aryCateInOrder(i)).ID Then
					Response.Write "<tr><td align=""center""><img width=""16"" src="""&ZC_BLOG_HOST&"ZB_SYSTEM/image/admin/arrow_turn_right.png"" alt="""" /></td>"
					Response.Write "<td>" & Categorys(aryCateInOrder(j)).ID & "</td>"
					Response.Write "<td>" & Categorys(aryCateInOrder(j)).Order & "</td>"
					Response.Write "<td><a href=""../../../catalog.asp?cate="& Categorys(aryCateInOrder(j)).ID &"""  target=""_blank"">&nbsp;┄&nbsp;" & Categorys(aryCateInOrder(j)).Name & "</a></td>"
					Response.Write "<td>" & Categorys(aryCateInOrder(j)).Alias & "</td>"
					rESPONSE.Write "<td>"&Replace(vbsunescape2(Categorys(aryCateInOrder(j)).Meta.GetValue("pagemeta")),vbcrlf,"<br/>")&"</td>"

					Response.Write "<td align=""center""><a href=""Edit.asp?act=2&id="& Categorys(aryCateInOrder(j)).ID &"""><img src="""&ZC_BLOG_HOST&"/ZB_SYSTEM/image/admin/folder_edit.png"" alt=""编辑"" title=""编辑"" width=""16"" /></a></td>"
					Response.Write "</tr>"
				End If

			Next

		End If

	Next
	End If

	Response.Write "</table>"


	PageMeta_ExportCategoryList=True

End Function

Function PageMeta_ExportUserList(intPage)

	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim intPageAll

	Call CheckParameter(intPage,"int",1)


	Call GetBlogHint()

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
	Response.Write "注意！一旦您修改了某个用户，该用户必须重新登录！"
	If (Not objRS.bof) And (Not objRS.eof) Then

		Response.Write "<table border=""1"" width=""100%"" cellspacing=""1"" cellpadding=""1"">"
		Response.Write "<tr><td>"& ZC_MSG076 &"</td><td>"&"</td><td>"& ZC_MSG001 &"</td><Td>Meta头 </td><td></td></tr>"

		For i=1 to objRS.PageSize

			Response.Write "<tr>"
			Response.Write "<td>" & objRS("mem_ID") & "</td>"
			Response.Write "<td>" & ZVA_User_Level_Name(objRS("mem_Level")) & "</td>"
			Response.Write "<td><a href=""../../../catalog.asp?auth="& objRS("mem_ID") &"""  target=""_blank"">" & objRS("mem_Name") & "</a></td>"
			dim mt
			set mt=new tmeta
			mt.loadstring=objRs("mem_Meta")
			Response.Write "<td>" & Replace(vbsunescape2(mt.GetValue("pagemeta")),vbcrlf,"<br/>") & "</td>"

			Response.Write "<td align=""center""><a href=""edit.asp?act=3&id="& objRS("mem_ID") &"""><img src="""&ZC_BLOG_HOST&"/ZB_SYSTEM/image/admin/folder_edit.png"" alt=""编辑"" title=""编辑"" width=""16"" /></a></td>"

			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

		Response.Write "</table>"

	End If

	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"list.asp?act=UserMng&page=")

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage
	Response.Write "</p></div>"

	objRS.Close
	Set objRS=Nothing

	PageMeta_ExportUserList=True

End Function

Function pAgEmEtA_eXpOrTtAgLiSt(InTpAgE)

	Dim i
	Dim objRS
	Dim strPage
	Dim intPageAll


	Call GetBlogHint()


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

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""1"" cellpadding=""1"">"
	Response.Write "<tr><td width=""5%"">"& ZC_MSG076 &"</td><td width=""25%"">"& ZC_MSG001 &"</td><td width=""40%"">Meta头</td><td width=""15%""></td></tr>"

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			Response.Write "<tr>"
			Response.Write "<td>" & objRS("tag_ID") & "</td>"
			Response.Write "<td>" & objRS("tag_Name") & "</td>"
			dim mt
			set mt=new tmeta
			mt.loadstring=objRs("tag_Meta")
			Response.Write "<td>" & Replace(vbsunescape2(mt.GetValue("pagemeta")),vbcrlf,"<br/>") & "</td>"

			Response.Write "<td align=""center""><a href=""Edit.asp?act=4&id="& objRS("tag_ID") &"""><img src="""&ZC_BLOG_HOST&"/ZB_SYSTEM/image/admin/folder_edit.png"" alt=""编辑"" title=""编辑"" width=""16"" /></a></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	Response.Write "</table>"

	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"list.asp?act=TagMng&amp;page=")

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage
	Response.Write "</p></div>"

	objRS.Close
	Set objRS=Nothing

	PAGEMETA_ExportTagList=True

end function

%>
