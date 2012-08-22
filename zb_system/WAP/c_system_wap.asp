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



'*********************************************************
' 目的：    主页
'*********************************************************
Function WapMain()
		
		'列表页模板暂不考虑
		Response.Write WapExport(Request("page"),Request("cate"),Request("auth"),Request("date"),Request("tags"),ZC_DISPLAY_MODE_ALL)
		Response.Write WapExportBar(Request("page"),intPageCount,Request("cate"),Request("auth"),Request("date"),Request("tags"),Request("q"))
		
		WapNav()
				
End Function


'*********************************************************
' 目的：    搜索
'*********************************************************
Function WapSearch()

		Response.Write WapExport(Request("page"),Request("cate"),Request("auth"),Request("date"),Request("tags"),ZC_DISPLAY_MODE_SEARCH)
		Response.Write WapExportBar(Request("page"),intPageCount,Request("cate"),Request("auth"),Request("date"),Request("tags"),Request("q"))
		
		WapNav()
				
End Function
'*********************************************************
' 目的：    导航
'*********************************************************
Function WapNav()
		Response.Write "<div class=""t2""></div>"		
		Response.Write "<div id=""nav"">"
		If BlogUser.Level>3 Then		
		Response.Write "<a href="""&WapUrlStr&"?act=Login"">"&ZC_MSG009&"</a><b>|</b>"
		End If		
		Response.Write "<a href="""&WapUrlStr&"?act=Com"">"&ZC_MSG027&"</a><b>|</b>"		
'		Response.Write "<p>5 <a  accesskey=""5""  href="""&WapUrlStr&"?act=Prev"">"&ZC_MSG032&"</a></p>"
		If Not ZC_DISPLAY_CATE_ALL_WAP Then
		Response.Write "8 <a href="""&WapUrlStr&"?act=Cate"">"&ZC_MSG214&"</a><b>|</b>"
		End If 
'		Response.Write "<p>7 <a href="""&WapUrlStr&"?act=Stat"">"&ZC_MSG029&"</a></p>"	
		If BlogUser.Level<=3 Then
		Response.Write "<a  href="""&WapUrlStr&"?act=AddArt"">"&ZC_MSG168&"</a><b>|</b>"	
		End If
'		Response.Write "<p>9 <a href="""&WapUrlStr&""">"&ZC_MSG213&"</a></p>"
		Response.Write "<a href="""&ZC_BLOG_HOST&""">电脑版</a>"		
		Response.Write "</div><div class=""adm"">" &WapCheckLogin
		Response.Write "</div>"
End Function


'*********************************************************
' 目的：    查看分类
'*********************************************************
Function WapCate()
	Dim Category
	Response.Write WapTitle(ZC_MSG214,"")
	Response.Write "<ul>"
		For Each Category in Categorys
			If IsObject(Category) Then
				Response.Write "<li>"&Category.ID&".<a href="""&WapUrlStr&"?act=Main&amp;cate="&Category.ID&""">"&TransferHTML(Category.Name,"[html-format]")&"</a>("&Category.Count&")</li>"
			End If
		Next
	Response.Write "</ul>"	
	WapNav()	
End Function


'*********************************************************
' 目的：    最新发表
'*********************************************************
Function WapPrev()
	Response.Write  WapTitle(ZC_MSG032,"")
	Response.Write  "<ul>" & LoadFromFile(BlogPath & "/INCLUDE/previous.asp","utf-8") & "</ul>"
	WapNav()
End Function


'*********************************************************
' 目的：    查看站点统计
'*********************************************************
Function WapStat()
	Response.Write  WapTitle(ZC_MSG026,"")
	Response.Write  "<ul>" & LoadFromFile(BlogPath & "/INCLUDE/statistics.asp","utf-8") & "</ul>"
	WapNav()
End Function


'*********************************************************
' 目的：    查看标题-页头
'*********************************************************
Public Function WapTitle(strCom,strBrowserTitle)

	If strBrowserTitle="" Then strBrowserTitle=strCom

	WapTitle = "<title>" &strBrowserTitle& "</title>"&vbCrLf
	WapTitle = WapTitle & "</head>"&vbCrLf
	WapTitle = WapTitle & "<body>"&vbCrLf
	WapTitle = WapTitle & "<h1>"&ZC_BLOG_TITLE&"</h1>"
	If ZC_DISPLAY_CATE_ALL_WAP Then 
		Dim Category
		WapTitle = WapTitle & "<div class=""h"">"
		    WapTitle = WapTitle & "<a href="""&WapUrlStr&""">"&ZC_MSG213&"</a><b>|</b>"
			For Each Category in Categorys
				If IsObject(Category) Then
					WapTitle = WapTitle & "<a href="""&WapUrlStr&"?act=Main&amp;cate="&Category.ID&""">"&TransferHTML(Category.Name,"[html-format]")&"</a><b>|</b>"
				End If
			Next
		WapTitle = Left(WapTitle, Len(WapTitle)-8) & "</div>"	
	End If 

	If IsEmpty(Request.QueryString("act")) Then 
	WapTitle = WapTitle & "<form action="""&WapUrlStr&""" method=""get"">"
    WapTitle = WapTitle & "    <div class=""srh"">"
	WapTitle = WapTitle & "	<input type=""hidden"" name=""act"" value=""Search"">"
    WapTitle = WapTitle & "    <input type=""text"" class=""i"" name=""q"" value="""" id=""q"">"
    WapTitle = WapTitle & "   <input type=""submit"" name=""submit"" value=""搜"">"
    WapTitle = WapTitle & "    </div>"
	WapTitle = WapTitle & "</form>"
	End If 

	If strCom<>"" Then WapTitle = WapTitle & "<h2 class=""t1"">"&strCom&"</h2>"

End Function



'*********************************************************
' 目的：    Wap页面地址
'*********************************************************
Function WapUrlStr()
   WapUrlStr=ZC_BLOG_HOST&ZC_FILENAME_WAP
End Function




'*********************************************************
' 目的：    登录页面
'*********************************************************
Function WapLogin()
	Dim User,Password

	User=Request.Form("username")
	Password=Request.Form("password")
	Call CheckParameter(User,"sql",Empty)
	Call CheckParameter(Password,"sql",Empty)

	If IsEmpty(User) OR IsEmpty(Password) Then
		Response.Write WapTitle(ZC_MSG009,"")
		Response.Write "    <form method=""post"" action="""&WapUrlStr&"?act=Login""> "
		Response.Write "	<p>"&ZC_MSG001&"：<input type=""text"" name=""username"" size=""12"" value="""" /></p>"
		Response.Write "	<p>"&ZC_MSG002&"：<input type=""password"" name=""password"" size=""12"" value="""" /></p>"
		Response.Write "	<p><input name=""btnSumbit"" type=""submit"" value="""&ZC_MSG087&"""/> </p> "
		Response.Write "	</form> "
	Else
		Response.Cookies("password")=md5(Password)
		Response.Cookies("password").Expires=Date+365
		session(ZC_BLOG_CLSID&"password")=md5(Password)
		Response.Cookies("username")=User
		Response.Cookies("username").Expires=Date+365
		session(ZC_BLOG_CLSID&"username")=User
		

		If BlogUser.Verify=False Then
			Call ShowError(8)
		Else
			Response.Write WapMain()
		End If

	End If

End Function



'*********************************************************
' 目的：    检查登录
'*********************************************************
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
		s=s&" <a href="""&WapUrlStr&"?act=Logout"">"&ZC_MSG020&"</a>"
	Else
		s=""
	End If

	If s<>"" Then 
		WapCheckLogin=s
	Else
		WapCheckLogin=""
	End If
		   
End Function



'*********************************************************
' 目的：    退出登录
'*********************************************************
Function WapLogout()

	Response.Cookies("username")=""
	Response.Cookies("password")=""
	session(ZC_BLOG_CLSID&"password")=""    
	session(ZC_BLOG_CLSID&"username")=""
	Response.Cookies("username")=Empty
	Response.Cookies("password")=Empty
	session(ZC_BLOG_CLSID&"password")=Empty    
	session(ZC_BLOG_CLSID&"username")=Empty
	
	Response.Write WapTitle(ZC_MSG020,"")
	Response.Redirect Request.ServerVariables("Http_Referer")

End Function



'*********************************************************
' 目的：    版权声明
'*********************************************************
Function WapCopyRight()

	WapCopyRight=vbsunescape(Request.Cookies("username"))

End Function




'*********************************************************
' 目的：    删除文章
'*********************************************************
Function WapDelArt()
	Dim ID,T
	ID=Request.QueryString("id")
	T=Request.QueryString("t")
	Response.Write WapTitle(ZC_MSG063&ZC_MSG048&" › "&T,"")
	'加入确认
	If Request.QueryString("con")="Y" Then 
		If DelArticle(Request.QueryString("id")) Then
			Call MakeBlogReBuild_Core()
			Response.Write "<p class=""n"">"&ZC_MSG266&"</p>"	
			Response.Write "<p class=""s""><a href="""&WapUrlStr&""">"&ZC_MSG213&"</a></p>"
		End if 
	Else 
		Dim strYUrl
		strYUrl=WapUrlStr &"?act=DelArt&amp;id="&ID&"&amp;con=Y"
		Response.Write "<p class=""s""><a href="""&strYUrl&""">确定</a> | <a href=""javascript:history.go(-1)"">取消</a></p>"
	End If 	

End Function



'*********************************************************
' 目的：    删除评论
'*********************************************************
Function WapDelCom()
    Dim ID,LOG_ID
	ID=Request.QueryString("id")
	LOG_ID=Request.QueryString("log_id")
	Response.Write WapTitle(ZC_MSG063&ZC_MSG013,"")
	'加入确认
	If Request.QueryString("con")="Y" Then 
		Call DelComment(ID,LOG_ID)
	'	Call MakeBlogReBuild_Core()
		Response.Write "<p class=""n"">"&ZC_MSG266&"</p>"	
		Response.Write "<p class=""s""><a href=""javascript:history.go(-2)"">"&ZC_MSG065&"</a></p>"
	Else 
		Dim strYUrl
		strYUrl=WapUrlStr &"?act=DelCom&amp;id="&ID&"&amp;log_id="&LOG_ID&"&amp;con=Y"
		Response.Write "<p class=""s""><a href="""&strYUrl&""">确定</a> | <a href=""javascript:history.go(-1)"">取消</a></p>"
	End If 
End Function




'*********************************************************
' 目的：    新建文章（编辑）
'*********************************************************
Function WapEdtArt()

	Dim Log_ID

	Response.Write WapTitle(ZC_MSG168,"")

	Response.Write "<form method=""post""  action="""&WapUrlStr&"?act=PostArt&amp;inpId="&Log_ID&""" >"
	Response.Write "<input type=""hidden"" name=""edtID"" value=""0"">"
	Response.Write "<select name=""edtAuthorID"" style=""display:none"">"
	Response.Write "<option value="""&BlogUser.ID&""">"&TransferHTML(BlogUser.Name,"[html-format]")&"</option>"
	Response.Write "</select>"	
	Response.Write "<p>"&ZC_MSG060&"：<input type=""text"" name=""edtTitle""  class=""i"" value=""""/></p>"
	Response.Write "<p>"&ZC_MSG012&"：<select name=""edtCateID"">"

		Dim Category
		For Each Category in Categorys
			If IsObject(Category) Then
				Response.Write "<option value="""&Category.ID&""">"&TransferHTML(Category.Name,"[html-format]")&"</option>"
			End If
		Next
	Response.Write "</select> ｜ "

	Response.Write "<select name=""edtLevel"">"
		Dim i
		For i=Ubound(ZVA_Article_Level_Name) to 1 step -1
			Response.Write "<option value="""& i &""">"& ZVA_Article_Level_Name(i) &"</option>"
		Next
	Response.Write "</select></p>"

	Response.Write "<input type=""hidden"" name=""edtYear"" value="""&Year(GetTime(now()))&""" /><input name=""edtMonth"" type=""hidden"" value="""&Month(GetTime(now()))&""" /><input name=""edtDay"" type=""hidden"" value="""&Day(GetTime(now()))&""" /><input name=""edtTime"" type=""hidden"" value="""& Hour(GetTime(now()))&":"&Minute(GetTime(now()))&":"&Second(GetTime(now()))&""" />"
	
	Response.Write "<p>"&ZC_MSG138&"：<input name=""edtTag""  class=""i""  maxlength=""100"" value="""" /></p>"

	Response.Write "<p>"&ZC_MSG147&"：<input name=""edtAlias"" class=""i""  maxlength=""100"" value="""" /></p>"
	Response.Write "<p>"&ZC_MSG055&"："
	Response.Write "<textarea   name=""txaContent""  class=""i"" maxlength=""5000""  rows=""3""></textarea></p>"
	Response.Write "<input  name=""txaIntro"" type=""hidden""  value="""" />"
	Response.Write "<p><input  type=""submit"" value="&ZC_MSG087&" /><span class=""stamp""><a href=""javascript:history.go(-1)"">"&ZC_MSG065&"</a></span></p>"

	Response.Write "</form>"
	
End Function


'*********************************************************
' 目的：    文章发表
'*********************************************************
Function WapPostArt()
	If PostArticle() Then
		Call MakeBlogReBuild_Core()
		Response.Write "<p class=""n"">"&ZC_MSG266&"</p>"	
		Call WapMain()
	End If
End Function




'*********************************************************
' 目的：    添加评论（编辑）
'*********************************************************
Function WapAddCom(PostType)

	If ZC_WAPCOMMENT_ENABLE=False Then Call ShowError(40): Exit Function
	
	Dim log_ID,Author,Content,Email,HomePage
	log_ID=Request("inpId")
'	Author=Request.Form("inpName")
'	Content=Request.Form("inpArticle")
'	Email=Request.Form("inpEmail")
'	HomePage=Request.Form("inpHomePage")

	Call CheckParameter(log_ID,"int",0)
	If log_ID=0 Then Call ShowError(3): Exit Function

	Dim Article
	If log_ID<>0 Then
		Set Article=New TArticle
		If Article.LoadInfoByID(log_ID) Then
			If Article.Level=1 Then Response.Write WapTitle(ZVA_Article_Level_Name(1),"")&ZVA_ErrorMsg(9):Exit Function
			If Article.Level=2 Then
				If Not CheckRights("Root") Then
					If (Article.AuthorID<>BlogUser.ID) Then Response.Write WapTitle(ZVA_Article_Level_Name(2),"")&ZVA_ErrorMsg(6):Exit Function
				End If
			End If
		End If
	End If

	Response.Write WapTitle(Article.Title &" › "& ZC_MSG024,"")

    If PostType<>0 Then
    Response.Write "<p class=""n"">"&ZVA_ErrorMsg(PostType)&"</p>"
	End If
	
	Response.Write "	<form  method=""post"" action="""&WapUrlStr&"?act=PostCom&amp;inpId="&Log_ID&""" > "
	If (PostType<>31) And (BlogUser.Level<=3) Then
		Response.Write "	<p>"&ZC_MSG001&"："&BlogUser.Name&"<input  type=""hidden"" name=""inpName"" value="""&BlogUser.Name&""" maxlength="""&ZC_USERNAME_MAX&"""/></p>"
		Response.Write "	<input type=""hidden"" name=""inpEmail"" value="""&BlogUser.Email&""" maxlength="""&ZC_EMAIL_MAX&"""  /> "
		Response.Write "	<input type=""hidden"" name=""inpHomePage"" value="""&BlogUser.HomePage&""" maxlength="""&ZC_HOMEPAGE_MAX&"""  />"	
	Else
		Response.Write "	<p>"&ZC_MSG001&"：<input type=""text"" name=""inpName"" value="""" maxlength="""&ZC_USERNAME_MAX&"""/></p>"
		If PostType=6 Then
		Response.Write "	<p>"&ZC_MSG002&"：<input type=""password""  name=""inpPass""  value="""" maxlength="""&ZC_PASSWORD_MAX&"""/></p>"
		End If
		If Request("m")="Y" Then 
			Response.Write "	<p>网站：<input type=""text"" name=""inpHomePage"" value="""" maxlength="""&ZC_HOMEPAGE_MAX&"""  /></p> "			
			Response.Write "	<p>"&ZC_MSG053&"：<input type=""text"" name=""inpEmail"" value="""" maxlength="""&ZC_EMAIL_MAX&"""  /></p> "

		Else 
			Response.Write "	<p><a class=""a"" href="""&WapUrlStr&"?act=AddCom&amp;inpId="&log_ID&"&amp;m=Y"">更多选项</a></p>"
		End If 
		Response.Write "	<input type=""hidden"" name=""inpEmail"" value="""" maxlength="""&ZC_EMAIL_MAX&"""  /></p> "
		Response.Write "	<input type=""hidden"" name=""inpHomePage"" value="""" maxlength="""&ZC_HOMEPAGE_MAX&"""  /></p> "
	End If
	Response.Write "	<p><textarea name=""txaArticle"" class=""i"" maxlength="""&ZC_CONTENT_MAX&""" rows=""3"" ></textarea></p> "
	Response.Write "	<p><input name=""btnSumbit"" type=""submit"" value="""&ZC_MSG087&"""/> <span class=""stamp""><a href=""javascript:history.go(-1)"">"&ZC_MSG065&"</a></span></p> "
	Response.Write "	</form> "
	
End Function



'*********************************************************
' 目的：    评论发表	******有待修改******
'*********************************************************
Function WapPostCom()

	If ZC_WAPCOMMENT_ENABLE=False Then Call ShowError(40): Exit Function

	'PostComment(strKey,intRevertCommentID)
	If Not IsEmpty(Request.Form("inpPass")) Then
		Response.Cookies("password")=md5(Request.Form("inpPass"))
		session(ZC_BLOG_CLSID&"password")=md5(Request.Form("inpPass"))
	    Response.Cookies("username")=Request.Form("inpName")
		session(ZC_BLOG_CLSID&"username")=Request.Form("inpName")
		Call WapCheckLogin
	End IF

	Dim objComment
	Dim objArticle

	Set objComment=New TComment
	Set objArticle=New TArticle

	objComment.log_ID=Request("inpID")
	objComment.AuthorID=BlogUser.ID

	objComment.ParentID=Request("ParentID")

	objComment.Author=Request.Form("inpName")
	objComment.Content=Request.Form("txaArticle")
	objComment.Email=Request.Form("inpEmail")
	objComment.HomePage=Request.Form("inpHomePage")

'	Call CheckParameter(ParentID,"int",0)
	If Not CheckRegExp(objComment.Author,"[username]") Then Call  WapAddCom(15) :Exit Function
	
	IF Len(objComment.Content)=0 Or Len(objComment.Content)>ZC_CONTENT_MAX Then
		Call  WapAddCom(46) :Exit Function
	End If

	IF Len(objComment.Email)>0 Then
		If Not CheckRegExp(objComment.Email,"[email]") Then Call  WapAddCom(29) :Exit Function
	End If

	IF Len(objComment.HomePage)>0 Then
		If InStr(objComment.HomePage,"http")=0 Then objComment.HomePage="http://" & objComment.HomePage
		If Not CheckRegExp(objComment.HomePage,"[homepage]") Then Call WapAddCom(30) :Exit Function
	End If

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

	Response.Write WapTitle(ZC_MSG024,"")&"<p class=""n"">"&ZC_MSG266&"</p>"


    Response.Write "<a href="""&WapUrlStr&"?act=View&amp;id="&objComment.log_ID&""">"&ZC_MSG065&ZC_MSG048&"</a>"
	Response.Redirect WapUrlStr &"?act=Com&id="&objComment.log_ID
'   Response.Write " | <a href="""&WapUrlStr&"?act=Com&amp;id="&objComment.log_ID&""">"&ZC_MSG212&"</a>"

	Set objComment=Nothing

	
End Function



'*********************************************************
' 目的：    查看评论   ******有待修改******
'*********************************************************
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
				If Article.Level=1 Then Response.Write WapTitle(ZVA_Article_Level_Name(1),"")&ZVA_ErrorMsg(9):Exit Function
				If Article.Level=2 Then
					If Not CheckRights("Root") Then
						If (Article.AuthorID<>BlogUser.ID) Then Response.Write WapTitle(ZVA_Article_Level_Name(2),"")&ZVA_ErrorMsg(6):Exit Function
					End If
				End If
			End If
				Response.Write WapTitle(Article.title&"›"&ZC_MSG013,"")
		End If

		Dim objRS
		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn

		objRS.Source="SELECT [comm_ID],[log_ID],[comm_AuthorID],[comm_Author],[comm_Content],[comm_Email],[comm_HomePage],[comm_PostTime],[comm_IP],[comm_Agent],[comm_Reply],[comm_LastReplyIP],[comm_LastReplyTime],[comm_ParentID],[comm_IsCheck],[comm_Meta] FROM [blog_Comment] WHERE ([blog_Comment].[log_ID]=" & log_ID &" )"
		objRS.Open()


		If (Not objRS.bof) And (Not objRS.eof) Then
		Response.Write "<ul>"		
		Dim strCTemplate,ComRecordCount
'		strCTemplate=Application(ZC_BLOG_CLSID & "TEMPLATE_WAP_ARTICLE_COMMENT")
		strCTemplate=GetTemplate("TEMPLATE_WAP_ARTICLE_COMMENT")		

			objRS.PageSize = ZC_COMMENT_COUNT_WAP
			intPageCount=objRS.PageCount
			ComRecordCount=objRS.RecordCount
			objRS.AbsolutePage = CurrentPage

			For i=1 To objRS.PageSize
					Dim objComment
					Set objComment=New TComment
					If objComment.LoadInfoByArray(Array(objRS("comm_ID"),objRS("log_ID"),objRS("comm_AuthorID"),objRS("comm_Author"),objRS("comm_Content"),objRS("comm_Email"),objRS("comm_HomePage"),objRS("comm_PostTime"),objRS("comm_IP"),objRS("comm_Agent"),objRS("comm_Reply"),objRS("comm_LastReplyIP"),objRS("comm_LastReplyTime"),objRS("comm_ParentID"),objRS("comm_IsCheck"),objRs("comm_Meta"))) Then
					Dim strC_Count
					strC_Count=ComRecordCount-((CurrentPage-1)*ZC_COMMENT_COUNT_WAP+i)+1

					ReDim Preserve aryStrC(i)
					aryStrC(i)=strCTemplate
					aryStrC(i)=Replace(aryStrC(i),"<#ZC_FILENAME_WAP#>",WapUrlStr)
					aryStrC(i)=Replace(aryStrC(i),"<#article/id#>",objRS("log_ID"))
					aryStrC(i)=Replace(aryStrC(i),"<#article/title#>",Article.title)
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

					TemplateTagsDic.Item("BlogTitle")=ZC_BLOG_TITLE

					aryTemplateTagsName=TemplateTagsDic.Keys
					aryTemplateTagsValue=TemplateTagsDic.Items

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
			Response.Write "<span class=""t"">"& ZC_MSG256
			If ZC_WAPCOMMENT_ENABLE Then 
				Response.Write " | <a href="""& WapUrlStr &"?act=AddCom&amp;inpId="&log_ID&""">"&ZC_MSG024&"</a></span>"
			Else 
				Response.Write "</span>"
			End If 
			Exit Function

		End If

		objRS.Close()
		Set objRS=Nothing
		
		Dim strC
		strC=Join(aryStrC)

		Dim a,b
		Dim PageBar
		PageBar=""
		
		If ZC_DISPLAY_PAGEBAR_ALL_WAP Then
			If intPageCount>ZC_PAGEBAR_COUNT_WAP Then
				a=CurrentPage-Cint((ZC_COMMENT_PAGEBAR_COUNT_WAP-1)/2)
				b=CurrentPage+ZC_COMMENT_PAGEBAR_COUNT_WAP-Cint((ZC_COMMENT_PAGEBAR_COUNT_WAP-1)/2)-1
				If a<=1 Then 
					a=1:b=ZC_COMMENT_PAGEBAR_COUNT_WAP
				End If
				If b>=intPageCount Then 
					b=intPageCount:a=intPageCount-ZC_COMMENT_PAGEBAR_COUNT_WAP+1
				End If
			Else
				a=1:b=intPageCount
			End If
			PageBar=" <a href="""&WapUrlStr&"?act=Com&amp;id="&log_ID&"&amp;Page=1"">[&lt;]</a> "			
			For i=a to b 		
				If i=CurrentPage Then
				PageBar=PageBar&" "&i&" "
				Else
				PageBar=PageBar&" <a href="""&WapUrlStr&"?act=Com&amp;id="&log_ID&"&amp;Page="&i&""">["&i&"]</a> "
				End If
			Next
			PageBar=PageBar&" <a href="""&WapUrlStr&"?act=Com&amp;id="&log_ID&"&amp;Page="&intPageCount&""">[&gt;]</a> "
		Else 
			If CurrentPage>1 Then
				If CurrentPage<intPageCount Then
					PageBar=PageBar&"<a href="""&WapUrlStr&"?act=Com&amp;id="&log_ID&"&amp;Page="&CurrentPage+1&""">下一页</a> | "
				End If
				PageBar=PageBar&"<a href="""&WapUrlStr&"?act=Com&amp;id="&log_ID&"&amp;Page="&CurrentPage-1&""">上一页</a> | "&CurrentPage&"/"&intPageCount
			ElseIf  (CurrentPage mod intPageCount)<>0 Then
				If PageBar<>"" Then PageBar=PageBar&" | "
				PageBar=PageBar&"<a href="""&WapUrlStr&"?act=Com&amp;id="&log_ID&"&amp;Page="&CurrentPage+1&""">下一页</a> | "&CurrentPage&"/"&intPageCount
			Else
				PageBar=""
			End If
		End If 

		strC=strC&"</ul><div class=""a"">"&PageBar&"</div>"

		If log_ID<>0 Then strC=strC&"<p class=""t""><a href="""& WapUrlStr &"?act=AddCom&amp;inpId="&log_ID&""">"&ZC_MSG024&"</a></p>"
		
		Response.Write strC

		WapNav()

End Function



'*********************************************************
' 目的：    查看文章
'*********************************************************
Function WapView()
	Dim Article,ZC_SINGLE_START,CurrentPage,i,log_ID
	CurrentPage=Request.QueryString("page")
	log_ID=Request.QueryString("id")
	Call CheckParameter(CurrentPage,"int",1)
	Call CheckParameter(log_ID,"int",0)
	
	If log_ID=0 Then Call ShowError(3) : Exit Function

	Set Article=New TArticle
	If Article.LoadInfoByID(log_ID) Then

			Article.Template="WAP_SINGLE"

			If Article.Level=1 Then Response.Write WapTitle(ZVA_Article_Level_Name(1),"")&ZVA_ErrorMsg(9):Exit Function
			If Article.Level=2 Then
				If Not CheckRights("Root") Then
					If (Article.AuthorID<>BlogUser.ID) Then Response.Write WapTitle(ZVA_Article_Level_Name(2),"")&ZVA_ErrorMsg(6):Exit Function
				End If
			End If

	        Response.Write WapTitle(Article.Title,"")
			Dim ArticleContent,PageCount,PageBar
			ArticleContent=Article.Content
			If ZC_DISPLAY_MODE_ALL_WAP Then 
				ArticleContent=TransferHTML(UBBCode(ArticleContent,"[face][link][autolink][font][code][image][typeset][media][flash][key][upload]"),"[html-japan][vbCrlf][upload]")
				ArticleContent=TransferHTML(ArticleContent,"[closehtml]")
			Else 
				PageCount = Int(Len(ArticleContent)/ZC_SINGLE_SIZE_WAP) + 1
				ZC_SINGLE_START=Cint((CurrentPage-1)*ZC_SINGLE_SIZE_WAP+1)
				If ZC_SINGLE_START<1 Then ZC_SINGLE_START=1
				ArticleContent=TransferHTML(ArticleContent,"[html-format][wapnohtml][nbsp-br]")
				ArticleContent=Mid(ArticleContent,ZC_SINGLE_START,ZC_SINGLE_SIZE_WAP)
				ArticleContent=TransferHTML(ArticleContent,"[closehtml]")

				If CurrentPage>1 Then
					PageBar="<a href="""&WapUrlStr&"?act=View&amp;id="&log_ID&"&amp;Page="&CurrentPage-1&""">&laquo;上一页</a>"
					If CurrentPage<PageCount Then
						PageBar=PageBar&" | <a ref="""&WapUrlStr&"?act=View&amp;id="&log_ID&"&amp;Page="&CurrentPage+1&""">下一页&raquo;</a>"
					End IF
				ElseIf  CurrentPage<PageCount Then
					If PageBar<>"" Then PageBar=PageBar&" | "
					PageBar=PageBar&"<a href="""&WapUrlStr&"?act=View&amp;id="&log_ID&"&amp;Page="&CurrentPage+1&""">下一页&raquo;</a>"
				Else
					PageBar=""
				End If			
				
				ArticleContent=ArticleContent&"<p>"&PageBar&"</p>"
			End If 

			If Article.Export(ZC_DISPLAY_MODE_ALL) Then
				Article.Build
				Article.html=Replace(Article.html,"<#article/PageContent#>",ArticleContent)
				Article.html=Replace(Article.html,"<#ZC_FILENAME_WAP#>",WapUrlStr)
				Article.html=Replace(Article.html,"<#template:article_mutuality#>",WapRelateList(Article.ID,Article.Tag))

			    If BlogUser.Level<=3 Then
	
					Article.html=Replace(Article.html,"<#adbegin#>","")
					Article.html=Replace(Article.html,"<#adend#>","")
				Else
					Dim objRegExp
					Set objRegExp=New RegExp
					objRegExp.IgnoreCase =True
					objRegExp.Global=True
					objRegExp.Pattern="<#adbegin#>(.+)<#adend#>"
					Article.html= objRegExp.Replace(Article.html,"")
				End If
				Response.Write Article.html
			End If
			
	End If


End Function




'*********************************************************
' 目的：    相关文章
'*********************************************************
Function WapRelateList(intID,strTag)
	If (intID=0) Or Not ZC_WAP_MUTUALITY_LIMIT=0 Then Exit Function
	If strTag<>"" Then

	Dim strCC_Count,strCC_ID,strCC_Name,strCC_Url,strCC_PostTime,strCC_Title
	Dim strCC
	Dim i
	Dim j
	Dim objRS
	Dim strSQL

	Dim intWapMutualityLimit
	Dim strWapMutuality

	intWapMutualityLimit = ZC_WAP_MUTUALITY_LIMIT
	strWapMutuality = GetTemplate("TEMPLATE_WAP_ARTICLE_MUTUALITY")

'	Call Add_Action_Plugin("Action_Plugin_System_Initialize","Call Wap_addMutualityTemplate()")

	Dim strOutput
	strOutput=""

	Set objRS=Server.CreateObject("ADODB.Recordset")

	strSQL="SELECT top 10 [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Level]>2) AND [log_ID]<>"& intID &" "

	strSQL = strSQL & " AND ("

	Dim aryTAGs
	strTag=Replace(strTag,"}","")
	aryTAGs=Split(strTag,"{")

	                    For j = LBound(aryTAGs) To UBound(aryTAGs)
	                            If aryTAGs(j)<>"" Then
	                                    strSQL = strSQL & "([log_Tag] Like '%{"&aryTAGs(j)&"}%')"
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
  
		For i=1 To intWapMutualityLimit '相关文章数目
		Dim objArticle
		Set objArticle=New TArticle
		If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then

		    strCC_Count=strCC_Count+1
		    strCC_ID=objArticle.ID
		    strCC_Url=objArticle.Url
		    strCC_PostTime=FormatDateTime(objArticle.PostTime,vbLongDate)
		    strCC_Title=objArticle.Title

			strCC=strWapMutuality

		    strCC=Replace(strCC,"<#article/mutuality/id#>",strCC_ID) 
		    strCC=Replace(strCC,"<#article/mutuality/posttime#>",strCC_PostTime) 
		    strCC=Replace(strCC,"<#article/mutuality/name#>",strCC_Title) 
			strCC=Replace(strCC,"<#ZC_FILENAME_WAP#>",WapUrlStr)

			strOutput=strOutput & strCC

		end if

		Set objArticle=nothing
		objRS.MoveNext
		If objRS.eof Then Exit For
		Next

	End if

	objRS.Close()
	Set objRS=Nothing
	End If

	WapRelateList= strOutput 
	
End Function




'*********************************************************
' 目的：    查看文章列表
'*********************************************************
Function WapExport(intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType)

		Dim i,j,s,intWapCount
		Dim objRS
		Dim objArticle
		Dim q,Search

		Call CheckParameter(intPage,"int",1)
		Call CheckParameter(intCateId,"int",Empty)
		Call CheckParameter(intAuthorId,"int",Empty)
		Call CheckParameter(dtmYearMonth,"dtm",Empty)

		'添加搜索
		If intType=ZC_DISPLAY_MODE_SEARCH Then 
			q=TransferHTML(Request("q"),"[nohtml]")
			q=Trim(q)
			If Len(q)=0 Then Search=True:Exit Function
		'过滤SQL
			q=FilterSQL(q)
			intWapCount = ZC_SEARCH_COUNT
		Else 
			intWapCount = ZC_DISPLAY_COUNT_WAP
		End If 


		Dim Title
		Title=""

		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn
		objRS.Source="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_ID]>0) AND ([log_Level]>1) AND ([log_Type]="&ZC_POST_TYPE_ARTICLE&")"


		'添加搜索
		If intType=ZC_DISPLAY_MODE_SEARCH Then
		objRS.Source="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_ID]>0) AND ([log_Level]>2)"
		objRS.Source=objRS.Source & "AND(( [log_Title] like '%"&q&"%') OR ([log_Intro] LIKE '%"&q&"%') OR ([log_Content] LIKE '%"&q&"%'))  AND ([log_Type]="&ZC_POST_TYPE_ARTICLE&")"
		End If 

		If Not IsEmpty(intCateId) Then
			objRS.Source=objRS.Source & "AND([log_CateID]="&intCateId&")"
			'On Error Resume Next
			Title=Categorys(intCateId).Name
			Err.Clear
		End if
		If Not IsEmpty(intAuthorId) Then
			objRS.Source=objRS.Source & "AND([log_AuthorID]="&intAuthorId&")"
			'On Error Resume Next
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
				'Template_Calendar=Application(ZC_BLOG_CLSID & "CACHE_INCLUDE_CALENDAR")
				Template_Calendar=GetTemplate("CACHE_INCLUDE_CALENDAR")
			End If
			Application.UnLock

			Title=Year(dtmYearMonth) & " " & ZVA_Month(Month(dtmYearMonth))
		End If
		If Not IsEmpty(strTagsName) Then
			'On Error Resume Next
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


		Dim strDTitle
		If Title="" Then strDTitle=ZC_BLOG_TITLE

		objRS.Source=objRS.Source & "ORDER BY [log_PostTime] DESC,[log_ID] DESC"
		objRS.Open()

		'添加搜索
		If intType=ZC_DISPLAY_MODE_SEARCH Then
		s=Replace(Replace(ZC_MSG086,"%s","<b>" & TransferHTML(Replace(q,Chr(39)&Chr(39),Chr(39),1,-1,0),"[html-format]") & "</b>",vbTextCompare,1),"%s","<b>" & objRS.RecordCount & "</b>",1,-1,0)
		strDTitle=ZC_MSG158
		Title=s
		End If 

		If (Not objRS.bof) And (Not objRS.eof) Then
			objRS.PageSize = intWapCount
			intPageCount=objRS.PageCount
			objRS.AbsolutePage = intPage

			For i = 1 To objRS.PageSize

				ReDim Preserve aryArticleList(i)

				Set objArticle=New TArticle
				If objArticle.LoadInfoByArray(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17))) Then
					objArticle.SearchText=Request.QueryString("q")
					objArticle.Template="WAP_ARTICLE-MULTI"
					If objArticle.Export(intType)= True Then
						aryArticleList(i)=objArticle.html
					End If
				End If
				Set objArticle=Nothing

				objRS.MoveNext
				If objRS.EOF Then Exit For

			Next

		Else
			WapExport= WapTitle(Title,strDTitle) &"<p class=""t"">"& ZC_MSG256 &"</p>"
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

		TemplateTagsDic.Item("BlogTitle")=Title

		aryTemplateTagsName=TemplateTagsDic.Keys
		aryTemplateTagsValue=TemplateTagsDic.Items


		j=UBound(aryTemplateTagsName)

		For i=0 to j
			Template_Article_Multi=Replace(Template_Article_Multi,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
		Next

		Template_Article_Multi=Replace(Template_Article_Multi,"<#ZC_FILENAME_WAP#>",WapUrlStr)
		If BlogUser.Level<=3 Then
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
		
		WapExport= WapTitle(Title,strDTitle) & "<ul>"&Template_Article_Multi&"</ul>"

End Function


'*********************************************************
' 目的：    列表分页
'*********************************************************
Function WapExportBar(intNowPage,intAllPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,strQuestion)

		Dim i
		Dim a,b,c
		Dim t
		Dim strPageBar,Template_PageBar


		Call CheckParameter(intNowPage,"int",1)

		t=t & "?"
		
		If Not IsEmpty(intCateId) Then t=t & "act=Main&amp;cate=" & intCateId & "&amp;"
		If Not IsEmpty(dtmYearMonth) Then t=t & "act=Main&amp;date=" & Year(dtmYearMonth) & "-" & Month(dtmYearMonth) & "&amp;"
		If Not IsEmpty(intAuthorId) Then t=t & "act=Main&amp;auth=" & intAuthorId & "&amp;"
		If Not (strTagsName="") Then t=t & "act=Main&amp;tags=" & Server.URLEncode(strTagsName) & "&amp;"
		If Not (strQuestion="") Then t=t & "act=Search&amp;q=" & Server.URLEncode(strQuestion) & "&amp;"

		
		If intAllPage>0 Then			
			If ZC_DISPLAY_PAGEBAR_ALL_WAP  Then	
				If intAllPage>ZC_PAGEBAR_COUNT_WAP Then
					a=intNowPage-Cint((ZC_PAGEBAR_COUNT_WAP-1)/2)
					b=intNowPage+ZC_PAGEBAR_COUNT_WAP-Cint((ZC_PAGEBAR_COUNT_WAP-1)/2)-1
					If a<=1 Then 
						a=1:b=ZC_PAGEBAR_COUNT_WAP
					End If
					If b>=intAllPage Then 
						b=intAllPage:a=intAllPage-ZC_PAGEBAR_COUNT_WAP+1
					End If
				Else
					a=1:b=intAllPage
				End If
				strPageBar=" <a href="""&WapUrlStr & t &"page=1"">[&lt;]</a> "			
				For i=a to b 		
					If i=intNowPage Then
					strPageBar=strPageBar&" "&i&" "
					Else
					strPageBar=strPageBar&" <a href="""&WapUrlStr& t &"page="&i&""">["&i&"]</a> "
					End If
				Next
				strPageBar=strPageBar&" <a href="""&WapUrlStr& t &"page="&intPageCount&""">[&gt;]</a> "
				Template_PageBar="<p>"  & strPageBar & "</p>"	
			Else 
				If intNowPage>1 Then
					If intNowPage<intAllPage Then
						strPageBar=strPageBar&"<a href="""&WapUrlStr& t &"page="& intNowPage+1 &""">下一页</a> | "
					End If
					strPageBar=strPageBar&"<a href="""&WapUrlStr& t &"page="&intNowPage-1&""">上一页</a> | "&intNowPage&"/"&intAllPage
				ElseIf  (intNowPage mod intPageCount)<>0 Then
					If strPageBar<>"" Then strPageBar=strPageBar&" | "
					strPageBar=strPageBar&"<a href="""&WapUrlStr& t &"page="&intNowPage+1&""">下一页</a> | "&intNowPage&"/"&intAllPage
				Else
					strPageBar=""
				End If
				Template_PageBar="<p>" & strPageBar & "</p>"
			End If 
		End If

		WapExportBar=Template_PageBar

End Function




'*********************************************************
' 目的：    查看错误
'*********************************************************
Public Function WapError()
	Dim ID
	ID=Request.QueryString("id")
	If Not IsNumeric(ID) Then
		ID=0
	ElseIf CINT(ID)>Ubound(ZVA_ErrorMsg) Or CINT(ID)<0 Then
		ID=0
	End If
	Response.Write WapTitle(ZVA_ErrorMsg(ID),"") & "<p class=""n"">"&ZVA_ErrorMsg(ID)&" <span class=""stamp""><a href=""javascript:history.go(-1)"">"&ZC_MSG065&"</a></span></p>"
End Function

'*********************************************************

Function ShowError_WAP(id)
	Response.Redirect WapUrlStr&"?act=Err&id="&id
End Function

%>
