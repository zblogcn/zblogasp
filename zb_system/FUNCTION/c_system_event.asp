<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:
'// 程序版本:
'// 单元名称:    c_system_event.asp
'// 开始时间:    2005.02.11
'// 最后修改:
'// 备    注:
'///////////////////////////////////////////////////////////////////////////////





'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    用户登陆
'*********************************************************
Public Function Login()

	If CheckVerifyNumber(Request.Form("edtCheckOut"))=False Then Call ShowError(38)

	Login=BlogUser.Verify

End Function
'*********************************************************




'*********************************************************
' 目的：    用户退出
'*********************************************************
Public Function Logout()

	'Response.Cookies("username")=""
	'Response.Cookies("password")=""
	Response.Write "<script language=""JavaScript"" src=""script/common.js"" type=""text/javascript""></script>"
	Response.Write "<script language=""JavaScript"" type=""text/javascript"">"
	Response.Write "function SetCookie(sName, sValue,iExpireDays) {if (iExpireDays){var dExpire = new Date();dExpire.setTime(dExpire.getTime()+parseInt(iExpireDays*24*60*60*1000));document.cookie = sName + ""="" + escape(sValue) + ""; expires="" + dExpire.toGMTString();}else{document.cookie = sName + ""="" + escape(sValue) + ""; path=/"";	}}"
	Response.Write "SetCookie(""username"","""","""");"
	Response.Write "SetCookie(""password"","""","""");"
	Response.Write "window.location=""" & ZC_BLOG_HOST & """;"
	Response.Write "</script>"

	Logout=True

End Function
'*********************************************************




'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    文件上抟
'*********************************************************
Function UploadFile(bolAutoName,bolReload)

	Dim objUpLoadFile
	Set objUpLoadFile=New TUpLoadFile

	objUpLoadFile.AuthorID=BlogUser.ID

	If bolReload=True Then
		ShowError_Custom="Response.Write ""<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'><html xmlns='http://www.w3.org/1999/xhtml' xml:lang='zh-CN' lang='zh-CN'><head>	<link rel='stylesheet' rev='stylesheet' href='CSS/admin.css' type='text/css' media='screen' /></head><body><form id='edit' name='edit' method='post'><p>" & ZC_MSG098 & ":" & """&ZVA_ErrorMsg(id)&""" & "&nbsp;&nbsp;<a href='cmd.asp?act=FileSnd'>" & ZC_MSG295 & "</a></p></form></body></html>"":Response.End"
	End If

	If objUpLoadFile.UpLoad(bolAutoName) Then

		UploadFile=True

		If bolReload=False Then Exit Function

		Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/><meta http-equiv=""Content-Language"" content=""zh-cn"" /><link rel=""stylesheet"" rev=""stylesheet"" href=""CSS/admin.css"" type=""text/css"" media=""screen"" /></head><body>"

		Response.Write "<form border=""1"" name=""edit"" id=""edit"" method=""post"" enctype=""multipart/form-data"" action="""& ZC_BLOG_HOST &"cmd.asp?act=FileSnd"">"
		Response.Write "<p><input type=""submit"" class=""button"" value="""& ZC_MSG237 &""" name=""B1"" />&nbsp;&nbsp;"& ZC_MSG236 &":"
		Response.Write ""& "<a href="""& objUpLoadFile.FullUrlPathName &""" target=""_blank"">"& objUpLoadFile.FullUrlPathName &"</a></p>"
		Response.Write "</form>"


		Dim strFileType
		Dim strFileName
		Dim strUPLOADDIR
		Dim strUPLOADDIR2
		If ZC_UPLOAD_DIRBYMONTH Then
			CreatDirectoryByCustomDirectory(ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now())))
			strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now())) & "/"
			strUPLOADDIR2 = "upload/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now())) & "/"
		Else
			strUPLOADDIR = ZC_UPLOAD_DIRECTORY & "/"
			strUPLOADDIR2 ="upload/"
		End If
		strFileType=LCase(objUpLoadFile.FileName)

		If (CheckRegExp(strFileType,"\.(jpeg|jpg|gif|png|bmp)$")=True) Then
			strFileName="[IMG]"&strUPLOADDIR2&objUpLoadFile.FileName&"[/IMG]"
		ElseIf (CheckRegExp(strFileType,"\.(swf)$")=True) Then
			strFileName="[FLASH=400,300,True]"&strUPLOADDIR2&objUpLoadFile.FileName&"[/FLASH]"
		ElseIf (CheckRegExp(strFileType,"\.(wmv|avi|asf)$")=True) Then
			strFileName="[WMV=400,300,True]"&strUPLOADDIR2&objUpLoadFile.FileName&"[/WMV]"
		ElseIf (CheckRegExp(strFileType,"\.(qt|mov)$")=True) Then
			strFileName="[QT=400,300,True]"&strUPLOADDIR2&objUpLoadFile.FileName&"[/QT]"
		ElseIf (CheckRegExp(strFileType,"\.(rm|rmvb|mpg|mpeg)$")=True) Then
			strFileName="[RM=400,300,True]"&strUPLOADDIR2&objUpLoadFile.FileName&"[/RM]"
		ElseIf (CheckRegExp(strFileType,"\.(wma)$")=True) Then
			strFileName="[WMA=True]"&strUPLOADDIR2&objUpLoadFile.FileName&"[/WMA]"
		ElseIf (CheckRegExp(strFileType,"\.(rm)$")=True) Then
			strFileName="[RA=True]"&strUPLOADDIR2&objUpLoadFile.FileName&"[/RA]"
		Else
			strFileName="[URL="&strUPLOADDIR2 & objUpLoadFile.FileName &"]"& objUpLoadFile.FileName &"[/URL]"
		End If

		'edit
		Response.Write "<script language=""Javascript"">try{parent.document.edit.txaContent.currPos.text+='"&strFileName&"';}catch(e){try{parent.document.edit.txaContent.value+='"&strFileName&"'}catch(e){}}</script>"
		'edit_widgeditor
		Response.Write "<script language=""Javascript"">try{parent.document.getElementById('txaContentWidgIframe').contentWindow.document.getElementsByTagName('body')[0].innerHTML+='"&strFileName&"'}catch(e){}</script>"
		'edit_fckeditor
		Response.Write "<script language=""Javascript"">try{parent.CKEDITOR.instances.txaContent.insertHtml('"&Replace(TransferHTML(UBBCode(strFileName,"[link][image][media][flash]"),"[upload]"),"'","\'")&"')}catch(e){}</script>"
		'edit_htmlarea
		Response.Write "<script language=""Javascript"">try{parent.document.getElementById('ta').parentNode.getElementsByTagName('iframe')[0].contentWindow.document.getElementsByTagName('body')[0].innerHTML+='"&strFileName&"'}catch(e){}</script>"
		'edit_tinymce
		Response.Write "<script language=""Javascript"">try{parent.document.getElementById('mce_editor_0').contentWindow.document.getElementsByTagName('body')[0].innerHTML+='"&strFileName&"'}catch(e){}</script>"
		'edit_ewebeditor
		Response.Write "<script language=""Javascript"">try{parent.document.getElementById('eWebEditor1').contentWindow.document.getElementsByTagName('body')[0].innerHTML+='"&strFileName&"'}catch(e){}</script>"
		
		Response.Write "</body></html>"

		'If bolReload=True Then Response.End

	Else

		If bolReload=True Then Response.Redirect "admin/admin.asp?act=FileSnd"

	End If

	Set objUpLoadFile=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Form of Send File
'*********************************************************
Function SendFile()

	Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/><meta http-equiv=""Content-Language"" content=""zh-cn"" /><link rel=""stylesheet"" rev=""stylesheet"" href=""CSS/admin.css"" type=""text/css"" media=""screen"" /><script src=""script/common.js"" type=""text/javascript""></script></head><body>"

	Response.Write "<form border=""1"" name=""edit"" id=""edit"" method=""post"" enctype=""multipart/form-data"" action="""& ZC_BLOG_HOST &"cmd.asp?act=FileUpload&reload=1"">"
	Response.Write "<p>"& ZC_MSG108 &": "
	Response.Write "<input type=""file"" id=""edtFileLoad"" name=""edtFileLoad"" size=""20"">  <input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" name=""B1"" onclick='document.getElementById(""edit"").action=document.getElementById(""edit"").action+""&filename=""+escape(document.getElementById(""edtFileLoad"").value)' /> <input class=""button"" type=""reset"" value="""& ZC_MSG088 &""" name=""B2"" />"
	Response.Write "&nbsp;<input type=""checkbox"" onclick='if(this.checked==true){document.getElementById(""edit"").action=document.getElementById(""edit"").action+""&autoname=1"";}else{document.getElementById(""edit"").action="""& ZC_BLOG_HOST &"cmd.asp?act=FileUpload&reload=1"";};SetCookie(""chkAutoFileName"",this.checked,365);' id=""chkAutoName"" id=""chkAutoName""/><label for=""chkAutoName"">"& ZC_MSG131 &"</label></p></form>"

	Response.Write "<script type=""text/javascript"">if(GetCookie(""chkAutoFileName"")==""true""){document.getElementById(""chkAutoName"").checked=true;document.getElementById(""edit"").action=document.getElementById(""edit"").action+""&autoname=1"";};</script></body></html>"

End Function
'*********************************************************






'*********************************************************
' 目的：     文件删除
'*********************************************************
Function DelFile(intID)

	Dim objUpLoadFile
	Set objUpLoadFile=New TUpLoadFile

	If objUpLoadFile.LoadInfoByID(intID) Then

		If (objUpLoadFile.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True) Then
			If objUpLoadFile.Del Then DelFile=True
		End If

	Else
		Exit Function
	End If

	Set objUpLoadFile=Nothing

End Function
'*********************************************************




'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    Post Article
'*********************************************************
Function PostArticle()

	Dim s,i,t,k
	Dim strTag


	If Request.Form("edtID")<>"0" Then
		Dim objTestArticle
		Set objTestArticle=New TArticle
		If objTestArticle.LoadInfobyID(Request.Form("edtID")) Then
			If Not((objTestArticle.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Exit Function
			strTag=objTestArticle.Tag
			objTestArticle.DelFile
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
	objArticle.TemplateName=Request.Form("edtTemplate")

	objArticle.Intro=Request.Form("txaIntro")

	Select Case LCase(Request.QueryString("type"))
	Case "ueditor"
		objArticle.Content=Request.Form("txaContent")
		If objArticle.Intro="" Then
			s=objArticle.Content
			If Len(s)>ZC_TB_EXCERPT_MAX Then
				i=InStr(s,vbCrlf)
				If i>0 Then
					t=Split(s,vblf)
					s=""
					For k=LBound(t) To UBound(t)
						s=s & t(k)
						If Len(s)>ZC_TB_EXCERPT_MAX Then Exit For
					Next
					s=Replace(s,vbCr,vbCrlf)
				End If
				s=s & ZC_MSG305
			End If
			s=TransferHTML(s,"[closehtml]")
			objArticle.Intro=s
		End If
	Case Else
		objArticle.Content=Request.Form("txaContent")
		If objArticle.Intro="" Then
			s=objArticle.Content
			If Len(s)>ZC_TB_EXCERPT_MAX Then
				i=InStr(s,vbCrlf)
				If i>0 Then
					t=Split(s,vblf)
					s=""
					For k=LBound(t) To UBound(t)
						s=s & t(k)
						If Len(s)>ZC_TB_EXCERPT_MAX Then Exit For
					Next
					s=Replace(s,vbCr,vbCrlf)
				End If
				s=s & ZC_MSG305
			End If
			s=TransferHTML(s,"[closehtml]")
			objArticle.Intro=s
		End If
	End Select

	'接口
	Call Filter_Plugin_PostArticle_Core(objArticle)

	If objArticle.Post Then
		Call ScanTagCount(strTag)
		Call ScanTagCount(objArticle.Tag)
		Call BuildArticle(objArticle.ID,True,True)
		PostArticle=True
		Call Filter_Plugin_PostArticle_Succeed(objArticle)
	End If

	Set objArticle=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Del Article
'*********************************************************
Function DelArticle(intID)

	Dim strTag

	If intID<>"" Then
		Dim objTestArticle
		Set objTestArticle=New TArticle
		If objTestArticle.LoadInfobyID(intID) Then
			If Not((objTestArticle.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Exit Function
			strTag=objTestArticle.Tag
		Else
			Call ShowError(9)
		End If
		Set objTestArticle=Nothing
	End If

	Dim objArticle
	Set objArticle=New TArticle

	If objArticle.LoadInfoByID(intID) Then

		Call ScanTagCount(objArticle.Tag)

		If objArticle.Del Then DelArticle=True

		Call ScanTagCount(strTag)

		Call BlogReBuild_Comments

		Dim objNavArticle
		Dim objRS
		Set objRS=objConn.Execute("SELECT TOP 1 [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND ([log_PostTime]<" & ZC_SQL_POUND_KEY & objArticle.PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] DESC")
		If (Not objRS.bof) And (Not objRS.eof) Then
			Call BuildArticle(objRS("log_ID"),False,False)
		End If
		Set objRS=Nothing
		Set objRS=objConn.Execute("SELECT TOP 1 [log_ID] FROM [blog_Article] WHERE ([log_Level]>2) AND ([log_PostTime]>" & ZC_SQL_POUND_KEY & objArticle.PostTime & ZC_SQL_POUND_KEY &") ORDER BY [log_PostTime] ASC")
		If (Not objRS.bof) And (Not objRS.eof) Then
			Call BuildArticle(objRS("log_ID"),False,False)
		End If
		Set objRS=Nothing

	End If

	Set objArticle=Nothing

End Function
'*********************************************************




'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    Post Category
'*********************************************************
Function PostCategory()

	Dim objCategory
	Set objCategory=New TCategory
	objCategory.ID=Request.Form("edtID")
	objCategory.Name=Request.Form("edtName")
	objCategory.Order=Request.Form("edtOrder")
	objCategory.ParentID=Request.Form("edtPareID")
	objCategory.Alias=Request.Form("edtAlias")
	objCategory.TemplateName=Request.Form("edtTemplate")

	'接口
	Call Filter_Plugin_PostCategory_Core(objCategory)


	If objCategory.Post Then

		PostCategory=True

		Call Filter_Plugin_PostCategory_Succeed(objCategory)

	End If

	Set objCategory=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Del Category
'*********************************************************
Function DelCategory(intID)

	Dim objCategory
	Set objCategory=New TCategory

	If objCategory.LoadInfobyID(intID) Then
		If objCategory.Del Then DelCategory=True
	End If

	Set objCategory=Nothing

End Function
'*********************************************************





'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    Post Comment
'*********************************************************
Function PostComment(strKey)

	If IsEmpty(Request.Form("inpAjax"))=False Then
		ShowError_Custom="Call RespondError(id,ZVA_ErrorMsg(id)):Response.End"
	End If

	If ZC_COMMENT_TURNOFF Then
		Call ShowError(40)
	End If

	If ZC_COMMENT_VERIFY_ENABLE Then
		If CheckVerifyNumber(Request.Form("inpVerify"))=False Then Call ShowError(38)
	End If

	Dim inpID,inpName,inpArticle,inpEmail,inpHomePage,inpParentID

	inpID=Request.Form("inpID")
	inpName=Request.Form("inpName")
	inpArticle=Request.Form("inpArticle")
	inpEmail=Request.Form("inpEmail")
	inpHomePage=Request.Form("inpHomePage")
	inpParentID=Request.Form("inpParentID")
	
	If Len(inpArticle)=0 Or Len(inpArticle)>ZC_CONTENT_MAX Then
		Call  ShowError(46)
	End If

	Dim objComment
	Dim objArticle
	Dim tmpCount

	Set objComment=New TComment
	'If clng(inpParentID)>0 Then
'		objComment.LoadInfoById(inpParentID)'
'		If objComment.ParentCount>ZC_MAXFL'OOR Then Exit Function'
'		tmpCount= objComment.ParentCount
'		Set objComment=Nothing	
'		Set  objComment=New TComment
'	Else
		tmpCount=-1
'	End If
	objComment.log_ID=inpID
	objComment.AuthorID=BlogUser.ID
	objComment.Author=inpName
	objComment.Content=inpArticle
	objComment.Email=inpEmail
	objComment.HomePage=inpHomePage
	objComment.ParentID=inpParentID
	objComment.ParentCount=tmpCount+1


	'接口
	Call Filter_Plugin_PostComment_Core(objComment)

	If objComment.AuthorID>0 Then
		objComment.Author=Users(objComment.AuthorID).Name
	End If

	If objComment.log_ID>0 Then
		Set objArticle=New TArticle
		If objArticle.LoadInfoByID(objComment.log_ID) Then
			If Not (strKey=objArticle.CommentKey) Then Call ShowError(43)
			If objArticle.Level<4 Then Call ShowError(44)
		End If
		Set objArticle=Nothing
	Else
		If Not (strKey=Left(MD5(ZC_BLOG_HOST & ZC_BLOG_CLSID & CStr(0) & CStr(Day(GetTime(Now())))),8)) Then Call ShowError(43)
	End If

	Dim objUser
	For Each objUser in Users
		If IsObject(objUser) Then
			If (UCase(objUser.Name)=UCase(objComment.Author)) And (objUser.ID<>objComment.AuthorID) Then Call ShowError(31)
		End If
	Next

	If objComment.Post Then
		If objComment.log_ID>0 Then
			Call BuildArticle(objComment.log_ID,False,True)
			BlogReBuild_Comments
		Else
			BlogReBuild_GuestComments
		End If
		PostComment=True
		Call Filter_Plugin_PostComment_Succeed(objComment)

	End if

	If IsEmpty(Request.Form("inpAjax"))=False Then
		Call ReturnAjaxComment(objComment)
		Call ClearGlobeCache
		Call LoadGlobeCache
	End If

	Set objComment=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Del Comment
'*********************************************************
Function DelComment(intID,intLog_ID)

	Dim objComment
	Dim objArticle

	Set objComment=New TComment
	Set objArticle=New TArticle

	If objComment.LoadInfobyID(intID) Then

		If objComment.log_ID>0 Then
			Dim objTestArticle
			Set objTestArticle=New TArticle
			If objTestArticle.LoadInfobyID(objComment.log_ID) Then
				If Not((objComment.AuthorID=BlogUser.ID) Or (objTestArticle.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Exit Function
			Else
				Call ShowError(9)
			End If
			Set objTestArticle=Nothing
		Else
			If Not ((objComment.log_ID=0) And (CheckRights("GuestBookMng")=True)) Then Exit Function
		End If
		DelChild objComment.ID
		If objComment.Del Then
			If objComment.log_ID>0 Then
				Call BuildArticle(objComment.log_ID,False,True)
				BlogReBuild_Comments
			Else
				BlogReBuild_GuestComments
			End If
			DelComment=True
		End If

	End If

	Set objComment=Nothing

End Function
'*********************************************************



'EventMark4
'*********************************************************
' 目的：    Revert Comment
'*********************************************************
Function RevertComment(strKey,intRevertCommentID)

	If IsEmpty(Request.Form("inpAjax"))=False Then
		ShowError_Custom="Call RespondError(id,ZVA_ErrorMsg(id)):Response.End"
	End If

	Call CheckParameter(intRevertCommentID,"int",0)

	If ZC_COMMENT_TURNOFF Then
		Call ShowError(40)
	End If

	If ZC_COMMENT_VERIFY_ENABLE Then
		If CheckVerifyNumber(Request.Form("inpVerify"))=False Then Call ShowError(38)
	End If

	Dim objComment
	Dim objArticle
	Dim inpID,inpName,inpArticle,inpEmail,inpHomePage,inpParentID

	Set objComment=New TComment
	inpID=Request.Form("inpID")
	inpName=Request.Form("inpName")
	inpArticle=Request.Form("inpArticle")
	inpEmail=Request.Form("inpEmail")
	inpHomePage=Request.Form("inpHomePage")
	inpParentID=intRevertCommentID

	If Len(inpArticle)=0 Or Len(inpArticle)>ZC_CONTENT_MAX Then
		Call  ShowError(46)
	End If
	Dim tmpCount
	Set objComment=New TComment
	If clng(inpParentID)>0 Then
		objComment.LoadInfoById(inpParentID)
		If objComment.ParentCount>ZC_MAXFLOOR-2 Then Call ShowError(52)
		tmpCount= objComment.ParentCount
		Set objComment=Nothing	
		Set  objComment=New TComment
	Else
		Call ShowError(53)
	End If
	
	objComment.log_ID=inpID
	objComment.AuthorID=BlogUser.ID
	objComment.Author=inpName
	objComment.Content=inpArticle
	objComment.Email=inpEmail
	objComment.HomePage=inpHomePage
	objComment.ParentID=inpParentID
	objComment.ParentCount=tmpCount+1

	If objComment.log_ID>0 Then
		Set objArticle=New TArticle
		If objArticle.LoadInfoByID(objComment.log_ID) Then
			If Not (strKey=objArticle.CommentKey) Then Call ShowError(43)
			If objArticle.Level<4 Then Call ShowError(44)
		Else
			Call ShowError(9)
		End If
		Set objArticle=Nothing
	Else
		If Not (strKey=Left(MD5(ZC_BLOG_HOST & ZC_BLOG_CLSID & CStr(0) & CStr(Day(GetTime(Now())))),8)) Then Call ShowError(43)
	End If

	'接口
	Call Filter_Plugin_PostComment_Core(objComment)

	If objComment.Post Then
		If objComment.log_ID>0 Then
			Call BuildArticle(objComment.log_ID,False,False)
			BlogReBuild_Comments
		Else
			BlogReBuild_GuestComments
		End If

		RevertComment=True

		Call Filter_Plugin_PostComment_Succeed(objComment)
	End if

	If IsEmpty(Request.Form("inpAjax"))=False Then
		objComment.LoadInfoById objComment.ParentID
		Call ReturnAjaxComment(objComment)
		Call ClearGlobeCache
		Call LoadGlobeCache
	End If

	Set objComment=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Save Comment
'*********************************************************
Function SaveComment(intID,intLog_ID)

	Dim objComment,objComment2
	Dim objArticle
	Dim inpParentID,tmpCount
	inpParentID=clng( Request.Form("intRepComment"))
	
	Set objComment=New TComment
	Set objComment2=New TComment
	

	
	If objComment.LoadInfoByID(intID)=True Then
	if inpParentID>0 And inpParentID<>clng(intID) then
		If objComment2.LoadInfoByID(inpParentID)=True Then
			If objComment2.ParentCount+1>ZC_MAXFLOOR Then Call SetBlogHint_Custom("x 超出了层数！"):SaveComment=True:Exit Function
			tmpCount=objComment2.ParentCount
			If objComment2.log_ID=cLng(intLog_ID) then
				objComment.ParentID=inpParentID
				objComment.ParentCount=tmpCount+1
			Else
				Call SetBlogHint_Custom("x 父评论和子评论不在同一篇文章!")
				SaveComment=True
				Exit Function
			End If
		End If
	else	
		If  inpParentID<>clng(intID) then objComment.parentid=0
	end if
		objComment.log_ID=intLog_ID
		objComment.Author=Request.Form("inpName")
		objComment.Email=Request.Form("inpEmail")
		objComment.HomePage=Request.Form("inpHomePage")
		objComment.Content=Request.Form("txaArticle") '& vbCrlf  & Replace(Replace(ZC_MSG273,"%s",BlogUser.Name,1,1),"%s",GetTime(Now()),1,1)
		objComment.Reply=Request.Form("txaReply")

	End If
	Set objComment2=Nothing
	If objComment.log_ID>0 Then
		Set objArticle=New TArticle
		If objArticle.LoadInfoByID(objComment.log_ID) Then
			If Not ((objArticle.AuthorID=BlogUser.ID) Or (objComment.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Exit Function
		End If
		Set objArticle=Nothing
	Else
		If Not ((objComment.log_ID=0) And (CheckRights("GuestBookMng")=True)) Then Exit Function
	End If

	If objComment.Post Then
		If objComment.log_ID>0 Then
			Call BuildArticle(objComment.log_ID,False,False)
			BlogReBuild_Comments
		Else
			BlogReBuild_GuestComments
		End If

		SaveComment=True

		Call Filter_Plugin_PostComment_Succeed(objComment)

	End if

	Set objComment=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Return Ajax Comment
'*********************************************************
Dim ReturnAjaxComment_aryTemplateTagsName
Dim ReturnAjaxComment_aryTemplateTagsValue

Function ReturnAjaxComment_Plugin(aryTemplateTagsName,aryTemplateTagsValue)
	ReturnAjaxComment_aryTemplateTagsName=aryTemplateTagsName
	ReturnAjaxComment_aryTemplateTagsValue=aryTemplateTagsValue
End Function
'Mark5
Function ReturnAjaxComment(objComment)

	Dim i,j
	i=0
	Dim objArticle

	If objComment.log_ID>0 Then
		'Filter_Plugin_TArticle_Export_TemplateTags
		Call Add_Filter_Plugin("Filter_Plugin_TArticle_Export_TemplateTags","ReturnAjaxComment_Plugin")
		Set objArticle=New TArticle
		If objArticle.LoadInfoByID(objComment.log_ID) Then
			Call objArticle.Export(ZC_DISPLAY_MODE_ALL)
			i=objArticle.CommNums
		End If
	Else
		'Filter_Plugin_TGuestBook_Export_TemplateTags
		Call Add_Filter_Plugin("Filter_Plugin_TGuestBook_Export_TemplateTags","ReturnAjaxComment_Plugin")
		Dim GuestBook
		Set GuestBook=New TGuestBook
		Call GuestBook.Export("")

		Dim objRS
		Set objRS=Server.CreateObject("ADODB.Recordset")
		objRS.CursorType = adOpenKeyset
		objRS.LockType = adLockReadOnly
		objRS.ActiveConnection=objConn
		objRS.Source=""
		objRS.Open("SELECT COUNT([comm_ID])AS allComment FROM [blog_Comment] WHERE [blog_Comment].[log_ID]=0")
		If (Not objRS.bof) And (Not objRS.eof) Then
			i=objRS("allComment")
		End If
		objRS.Close
		Set objRS=Nothing
	End If

	Dim strC
	strC=GetTemplate("TEMPLATE_B_ARTICLE_COMMENT")
	objComment.Count=objComment.Count+1
	strC=objComment.MakeTemplate(strC,True)

	strC=Replace(strC,"<#ZC_BLOG_HOST#>",ZC_BLOG_HOST)

	Dim aryTemplateTagsName2
	Dim aryTemplateTagsValue2

	aryTemplateTagsName2=TemplateTagsName
	aryTemplateTagsValue2=TemplateTagsValue

	j=UBound(aryTemplateTagsName2)

	For i=1 to j
		strC=Replace(strC,"<#" & aryTemplateTagsName2(i) & "#>",aryTemplateTagsValue2(i))
	Next

	j=UBound(ReturnAjaxComment_aryTemplateTagsName)
	For i=1 to j
		strC = Replace(strC,"<#" & ReturnAjaxComment_aryTemplateTagsName(i) & "#>", ReturnAjaxComment_aryTemplateTagsValue(i))

	Next

	strC= Replace(strC,vbCrLf,"")
	strC= Replace(strC,vbLf,"")
	strC= Replace(strC,vbTab,"")
	Call SETBLOGHINT_CUSTOM(STRC)
	Response.Write strC

	ReturnAjaxComment=True

End Function
'*********************************************************


'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    Post TrackBack
'*********************************************************
Function PostTrackBack(intID,strKey)

	Dim objTrackBack
	Dim objArticle

	Dim keys
	Dim i,j,k,b

	If ZC_TRACKBACK_TURNOFF Then
		Call RespondError(41,ZVA_ErrorMsg(41))
	End If

	If Len(strKey)=5 Then

		If CheckVerifyNumber(strKey)=False Then Call ShowError(43)

	ElseIf Len(strKey)=8 Then

		Set objArticle=New TArticle
		If objArticle.LoadInfoByID(intID) Then
			If Not (strKey=objArticle.TrackBackKey) Then Call RespondError(43)
			If objArticle.Level<4 Then Call RespondError(44)
		End If
		Set objArticle=Nothing

	Else
		Exit Function
	End If

	Set objTrackBack=New TTrackBack
	Set objArticle=New TArticle

	objTrackBack.log_ID=intID
	objTrackBack.URL=Request.Form("url")
	objTrackBack.Title=Request.Form("title")
	objTrackBack.Blog=Request.Form("blog_name")
	objTrackBack.Excerpt=Request.Form("excerpt")

	'接口
	Call Filter_Plugin_PostTrackBack_Core(objTrackBack)

	If objTrackBack.Post Then
		Call BuildArticle(objTrackBack.log_ID,False,True)
		BlogReBuild_TrackBacks
		PostTrackBack=True
		Call Filter_Plugin_PostTrackBack_Succeed(objTrackBack)
	End If

	Response.ContentType = "text/xml"
	Response.Clear
	Response.Write objTrackBack.TbXML

	Set objTrackBack=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Del TrackBack
'*********************************************************
Function DelTrackBack(intID,intLog_ID)

	Dim objTrackBack
	Dim objArticle

	Set objTrackBack=New TTrackBack
	Set objArticle=New TArticle

	If objTrackBack.LoadInfobyID(intID) Then

		Dim objTestArticle
		Set objTestArticle=New TArticle
		If objTestArticle.LoadInfobyID(objTrackBack.log_ID) Then
			If Not((objTestArticle.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Exit Function
		Else
			Call ShowError(9)
		End If
		Set objTestArticle=Nothing

		If objTrackBack.Del Then
			Call BuildArticle(objTrackBack.log_ID,False,True)
			BlogReBuild_TrackBacks
			DelTrackBack=True
		End If

	End If

	Set objTrackBack=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Send TrackBack
'*********************************************************
Function SendTrackBack()

	Dim objTrackBack
	Dim objArticle

	Set objTrackBack=New TTrackBack
	Set objArticle=New TArticle

	If objArticle.LoadInfobyID(Request.Form("edtID")) Then
		objTrackBack.URL=objArticle.Url
		objTrackBack.Title=objArticle.Title
		objTrackBack.Blog=ZC_BLOG_NAME
		objTrackBack.Excerpt=Left(objArticle.HtmlContent,250)
	Else
		Call ShowError(9)
	End If

	If objTrackBack.Send(Request.Form("edtTrackBack")) Then SendTrackBack=True
	Set objTrackBack=Nothing

End Function
'*********************************************************




'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    Edit User
'*********************************************************
Function EditUser()
	Dim objUser
	Set objUser=New TUser
	objUser.ID=Request.Form("edtID")
	objUser.Level=Request.Form("edtLevel")
	objUser.Name=Request.Form("edtName")
	objUser.PassWord=Request.Form("edtPassWord")
	objUser.Email=Request.Form("edtEmail")
	objUser.HomePage=Request.Form("edtHomePage")
	objUser.Alias=Request.Form("edtAlias")

	If Not((CInt(objUser.ID)=BlogUser.ID) Or (CheckRights("Root")=True)) Then Exit Function

	'接口
	Call Filter_Plugin_EditUser_Core(objUser)

	If objUser.Edit(BlogUser) Then
		Call Filter_Plugin_EditUser_Succeed(objUser)
		EditUser=True
	End IF

	Set objUser=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Del User
'*********************************************************
Function DelUser(intID)

	Dim objRS
	Dim objUser
	Dim objUpLoadFile

	Set objUser=New TUser
	objUser.ID=intID
	If objUser.Del(BlogUser) Then DelUser=True
	Set objUser=Nothing

End Function
'*********************************************************




'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    Blog ReBuild
'*********************************************************
Function MakeBlogReBuild()

	'plugin node
	bAction_Plugin_MakeBlogReBuild_Begin=False
	For Each sAction_Plugin_MakeBlogReBuild_Begin in Action_Plugin_MakeBlogReBuild_Begin
		If Not IsEmpty(sAction_Plugin_MakeBlogReBuild_Begin) Then Call Execute(sAction_Plugin_MakeBlogReBuild_Begin)
		If bAction_Plugin_MakeBlogReBuild_Begin=True Then Exit Function
	Next

	Call MakeBlogReBuild_Core()

	Call SetBlogHint(True,False,Empty)

	Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><link rel=""stylesheet"" rev=""stylesheet"" href=""CSS/admin.css"" type=""text/css"" media=""screen"" /></head><body>"

	Response.Write "<div id=""divMain""><div class=""Header"">" & ZC_MSG072 & "</div>"
	Response.Write "<div id=""divMain2"">"
	Call GetBlogHint()
	Response.Write "<form  name=""edit"" id=""edit"">"

	Response.Write "<p>" & ZC_MSG225 &"</p>"
	Response.Write "<p>" & Replace(ZC_MSG169,"%n",RunTime/1000)&"</p>"

	Response.Write "</form></div></div>"
	Response.Write "</body></html>"

	MakeBlogReBuild=True

	'plugin node
	bAction_Plugin_MakeBlogReBuild_End=False
	For Each sAction_Plugin_MakeBlogReBuild_End in Action_Plugin_MakeBlogReBuild_End
		If Not IsEmpty(sAction_Plugin_MakeBlogReBuild_End) Then Call Execute(sAction_Plugin_MakeBlogReBuild_End)
		If bAction_Plugin_MakeBlogReBuild_End=True Then Exit Function
	Next

End Function
'*********************************************************





'*********************************************************
' 目的：    All Files ReBuild
'*********************************************************
Function MakeFileReBuild()

	On Error Resume Next

	'plugin node
	bAction_Plugin_MakeFileReBuild_Begin=False
	For Each sAction_Plugin_MakeFileReBuild_Begin in Action_Plugin_MakeFileReBuild_Begin
		If Not IsEmpty(sAction_Plugin_MakeFileReBuild_Begin) Then Call Execute(sAction_Plugin_MakeFileReBuild_Begin)
		If bAction_Plugin_MakeFileReBuild_Begin=True Then Exit Function
	Next

	Dim intPage
	Dim intAllTime

	intPage=CInt(Request.QueryString("page"))
	intAllTime=CLng(Request.QueryString("all"))

	If intPage=0 Then
		Call MakeBlogReBuild_Core()
		intPage=1
		Response.Redirect ZC_BLOG_HOST&"zb_system/cmd.asp?act=FileReBuild&page="&intPage&"&all="&intAllTime
	End If

	Dim i,j

	Dim objRS
	Dim objArticle

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source="SELECT [log_ID] FROM [blog_Article] WHERE [log_Level]>1"
	objRS.Open()

	If (Not objRS.bof) And (Not objRS.eof) Then

		objRS.PageSize = ZC_REBUILD_FILE_COUNT

		If intPage>objRS.PageCount Then

			Call SetBlogHint(True,Empty,False)

			Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><link rel=""stylesheet"" rev=""stylesheet"" href=""CSS/admin.css"" type=""text/css"" media=""screen"" /></head><body>"

			Response.Write "<div id=""divMain""><div class=""Header"">" & ZC_MSG073 & "</div>"
			Response.Write "<div id=""divMain2"">"
			Call GetBlogHint()
			Response.Write "<form  name=""edit"" id=""edit"">"

			Response.Write "<p>" & ZC_MSG225 &"</p>"
			Response.Write "<p>" & Replace(ZC_MSG169,"%n",intAllTime/1000)&"</p>"

			Response.Write "</form></div></div>"
			Response.Write "</body></html>"

			Response.Cookies("FileReBuild_Step")=""
			Response.Cookies("FileReBuild_Step").Expires= (now()-1)

			MakeFileReBuild=True

			'plugin node
			bAction_Plugin_MakeFileReBuild_End=False
			For Each sAction_Plugin_MakeFileReBuild_End in Action_Plugin_MakeFileReBuild_End
				If Not IsEmpty(sAction_Plugin_MakeFileReBuild_End) Then Call Execute(sAction_Plugin_MakeFileReBuild_End)
				If bAction_Plugin_MakeFileReBuild_End=True Then Exit Function
			Next

			Exit Function

		End If

		objRS.AbsolutePage = intPage

		For i = 1 To ZC_REBUILD_FILE_COUNT

			Call BuildArticle(objRS("log_ID"),False,False)

			objRS.MoveNext
			If objRS.eof Then Exit For
		Next

		intAllTime=CLng(intAllTime)+RunTime

		Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/><meta http-equiv=""Content-Language"" content=""zh-cn"" /><meta http-equiv=""refresh"" content="""&ZC_REBUILD_FILE_INTERVAL&";URL=cmd.asp?act=FileReBuild&page="&intPage+1&"&all="&intAllTime&"""/><link rel=""stylesheet"" rev=""stylesheet"" href=""CSS/admin.css"" type=""text/css"" media=""screen"" /><title>"&ZC_MSG073&"</title></head><body>"

		Response.Write "<div id=""divMain""><div class=""Header"">" & ZC_MSG073 & "</div>"
		Response.Write "<div id=""divMain2"">"
		Response.Write "<form  name=""edit"" id=""edit"">"


		For j=1 To intPage
		Response.Write "<p>" &Replace(ZC_MSG227,"%n",j)&"</p>"
		Next

		Response.Write "<p>" &Replace(ZC_MSG152,"%n",ZC_REBUILD_FILE_INTERVAL)&"</p>"
		Response.Write "</form></div></div>"
		Response.Write "</body></html>"

		Response.Cookies("FileReBuild_Step")=intPage+1
		Response.Cookies("FileReBuild_Step").Expires= (now()+1)


	Else

		Call SetBlogHint(True,Empty,False)
		Response.Redirect "admin/admin.asp?act=AskFileReBuild"

	End If

	Err.Clear

End Function
'*********************************************************




'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    List User Rights
'*********************************************************
Function ListUser_Rights()

	Dim s
	Dim i
	Dim strAction
	Dim aryAction

	strAction="login|verify|logout|admin|cmt|tb|vrs|BlogReBuild|FileReBuild|ArticleMng|ArticleEdt|ArticlePst|ArticleDel|CategoryMng|CategoryPst|CategoryDel|CommentMng|CommentDel|CommentRev|TrackBackMng|TrackBackDel|TrackBackSnd|UserMng|UserEdt|UserCrt|UserDel|FileMng|FileUpload|FileDel|Search|TagMng|TagEdt|TagPst|TagDel|SettingMng|SettingSav|PlugInMng|rss|Root"

	aryAction=Split(strAction, "|")

	s=ZC_MSG019

	Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/><meta http-equiv=""Content-Language"" content=""zh-cn"" /><link rel=""stylesheet"" rev=""stylesheet"" href=""CSS/admin.css"" type=""text/css"" media=""screen"" /><title>"&ZC_MSG021&"</title></head><body>"

	Response.Write "<div id=""divMain""><div class=""Header"">" & ZC_MSG021 & "</div>"
	Response.Write "<div id=""divMain2""><form  name=""edit"" id=""edit""><P>"

	Response.Write ZC_MSG001 & ":" & BLogUser.Name & "<br/><br/>"
	Response.Write ZC_MSG249 & ":" & ZVA_User_Level_Name(BLogUser.Level) & "<br/><br/>"

	For i=LBound(aryAction) To UBound(aryAction)
		If Not CheckRights(aryAction(i)) Then s=Replace(s,"%s",":<font color=Red>fail</font>"&"<br/><br/>",1,1) Else s=Replace(s,"%s",":<font color=green>ok</font>"&"<br/><br/>",1,1)

	Next

	Response.Write s

	Response.Write "</p></form></div></div>"
	Response.Write "</body></html>"

	ListUser_Rights=True

End Function
'*********************************************************





'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    Save Blog Setting
'*********************************************************
Function SaveSetting()

	On Error Resume Next

	Dim i,j
	Dim s,t
	Dim strContent

	strContent=LoadFromFile(BlogPath & "zb_users/c_custom.asp","utf-8")

	Dim strZC_BLOG_HOST
	Dim strZC_BLOG_TITLE
	Dim strZC_BLOG_SUBTITLE
	Dim strZC_BLOG_NAME
	Dim strZC_BLOG_SUB_NAME
	Dim strZC_BLOG_CSS
	Dim strZC_BLOG_THEME
	Dim strZC_BLOG_COPYRIGHT
	Dim strZC_BLOG_MASTER


	strZC_BLOG_HOST=Request.Form("edtZC_BLOG_HOST")
	If Right(strZC_BLOG_HOST,1)<>"/" Then strZC_BLOG_HOST=strZC_BLOG_HOST & "/"
	If Left(strZC_BLOG_HOST,8)<>"https://" Then
		If Left(strZC_BLOG_HOST,7)<>"http://" Then strZC_BLOG_HOST="http://" & strZC_BLOG_HOST
	End If

	strZC_BLOG_TITLE=Request.Form("edtZC_BLOG_TITLE")

	strZC_BLOG_SUBTITLE=Request.Form("edtZC_BLOG_SUBTITLE")

	strZC_BLOG_NAME=Request.Form("edtZC_BLOG_NAME")

	strZC_BLOG_SUB_NAME=Request.Form("edtZC_BLOG_SUB_NAME")

	strZC_BLOG_CSS=Request.Form("edtZC_BLOG_CSS")

	strZC_BLOG_THEME=Request.Form("edtZC_BLOG_THEME")

	strZC_BLOG_COPYRIGHT=Replace(Replace(Request.Form("edtZC_BLOG_COPYRIGHT"),vbCr,""),vbLf,"")

	strZC_BLOG_MASTER=Request.Form("edtZC_BLOG_MASTER")

	Call ScanPluginToThemeFile(strZC_BLOG_CSS,strZC_BLOG_THEME)

	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_HOST",strZC_BLOG_HOST)
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_TITLE",strZC_BLOG_TITLE)
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_SUBTITLE",strZC_BLOG_SUBTITLE)
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_NAME",strZC_BLOG_NAME)
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_SUB_NAME",strZC_BLOG_SUB_NAME)
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_CSS",strZC_BLOG_CSS)
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_THEME",strZC_BLOG_THEME)
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_COPYRIGHT",strZC_BLOG_COPYRIGHT)
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_MASTER",strZC_BLOG_MASTER)

	If UCase(strZC_BLOG_HOST)<>UCase("""" & CStr(ZC_BLOG_HOST) & """") Then Call SetBlogHint(Empty,Empty,True)
	If UCase(strZC_BLOG_TITLE)<>UCase("""" & CStr(ZC_BLOG_TITLE) & """") Then Call SetBlogHint(Empty,Empty,True)
	If UCase(strZC_BLOG_SUBTITLE)<>UCase("""" & CStr(ZC_BLOG_SUBTITLE) & """") Then Call SetBlogHint(Empty,Empty,Empty)
	If UCase(strZC_BLOG_NAME)<>UCase("""" & CStr(ZC_BLOG_NAME) & """") Then Call SetBlogHint(Empty,Empty,True)
	If UCase(strZC_BLOG_SUB_NAME)<>UCase("""" & CStr(ZC_BLOG_SUB_NAME) & """") Then Call SetBlogHint(Empty,Empty,True)
	If UCase(strZC_BLOG_CSS)<>UCase("""" & CStr(ZC_BLOG_CSS) & """") Then Call SetBlogHint(Empty,True,Empty)
	If UCase(strZC_BLOG_THEME)<>UCase("""" & CStr(ZC_BLOG_THEME) & """") Then Call SetBlogHint(Empty,True,True)
	If UCase(strZC_BLOG_COPYRIGHT)<>UCase("""" & CStr(ZC_BLOG_COPYRIGHT) & """") Then Call SetBlogHint(Empty,Empty,True)
	If UCase(strZC_BLOG_MASTER)<>UCase("""" & CStr(ZC_BLOG_MASTER) & """") Then Call SetBlogHint(Empty,True,Empty)

	Call SaveToFile(BlogPath & "zb_users/c_custom.asp",strContent,"utf-8",False)

	strContent=LoadFromFile(BlogPath & "zb_users/c_option.asp","utf-8")

	Dim strZC_BLOG_CLSID
	strZC_BLOG_CLSID=Request.Form("edtZC_BLOG_CLSID")
	If CheckRegExp(strZC_BLOG_CLSID,"[guid]") Then
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_CLSID",strZC_BLOG_CLSID)
	If UCase(strZC_BLOG_CLSID)<>UCase("""" & CStr(ZC_BLOG_CLSID) & """") Then Call SetBlogHintWithCLSID(True,True,True,Replace(strZC_BLOG_CLSID,"""",""))
	End If

	Dim strZC_TIME_ZONE
	strZC_TIME_ZONE=Request.Form("edtZC_TIME_ZONE")
	Call SaveValueForSetting(strContent,True,"String","ZC_TIME_ZONE",strZC_TIME_ZONE)
	If UCase(strZC_TIME_ZONE)<>UCase("""" & CStr(ZC_TIME_ZONE) & """") Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_HOST_TIME_ZONE
	strZC_HOST_TIME_ZONE=Request.Form("edtZC_HOST_TIME_ZONE")
	Call SaveValueForSetting(strContent,True,"String","ZC_HOST_TIME_ZONE",strZC_HOST_TIME_ZONE)
	If UCase(strZC_HOST_TIME_ZONE)<>UCase("""" & CStr(ZC_HOST_TIME_ZONE) & """") Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_BLOG_LANGUAGE
	strZC_BLOG_LANGUAGE=Request.Form("edtZC_BLOG_LANGUAGE")
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_LANGUAGE",strZC_BLOG_LANGUAGE)
	If UCase(strZC_BLOG_LANGUAGE)<>UCase("""" & CStr(ZC_BLOG_LANGUAGE) & """") Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_UPDATE_INFO_URL
	strZC_UPDATE_INFO_URL=Request.Form("edtZC_UPDATE_INFO_URL")
	If (Not CheckRegExp(strZC_UPDATE_INFO_URL,"[homepage]")) And (strZC_UPDATE_INFO_URL<>"") Then strZC_UPDATE_INFO_URL="http://update.rainbowsoft.org/info/"
	Call SaveValueForSetting(strContent,True,"String","ZC_UPDATE_INFO_URL",strZC_UPDATE_INFO_URL)
	If UCase(strZC_UPDATE_INFO_URL)<>UCase("""" & CStr(ZC_UPDATE_INFO_URL) & """") Then Call SetBlogHint(Empty,Empty,Empty)

	Dim strZC_STATIC_TYPE
	strZC_STATIC_TYPE=Request.Form("edtZC_STATIC_TYPE")
	Call SaveValueForSetting(strContent,True,"String","ZC_STATIC_TYPE",strZC_STATIC_TYPE)
	If UCase(strZC_STATIC_TYPE)<>UCase("""" & CStr(ZC_STATIC_TYPE) & """") Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_STATIC_DIRECTORY
	strZC_STATIC_DIRECTORY=Request.Form("edtZC_STATIC_DIRECTORY")
	If LCase(Left(strZC_STATIC_DIRECTORY,3))="zb_" Then strZC_STATIC_DIRECTORY=Right(strZC_STATIC_DIRECTORY,Len(strZC_STATIC_DIRECTORY)-3)
	Call SaveValueForSetting(strContent,True,"String","ZC_STATIC_DIRECTORY",strZC_STATIC_DIRECTORY)
	If UCase(strZC_STATIC_DIRECTORY)<>UCase("""" & CStr(ZC_STATIC_DIRECTORY) & """") Then Call SetBlogHint(Empty,Empty,True)



	Dim strZC_BLOG_WEBEDIT
	strZC_BLOG_WEBEDIT=Request.Form("edtZC_BLOG_WEBEDIT")
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_WEBEDIT",strZC_BLOG_WEBEDIT)
	If UCase(strZC_BLOG_WEBEDIT)<>UCase("""" & CStr(ZC_BLOG_WEBEDIT) & """") Then Call SetBlogHint(Empty,Empty,Empty)

	Dim strZC_REBUILD_FILE_COUNT
	strZC_REBUILD_FILE_COUNT=Request.Form("edtZC_REBUILD_FILE_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_REBUILD_FILE_COUNT",strZC_REBUILD_FILE_COUNT)
	If UCase(strZC_REBUILD_FILE_COUNT)<>UCase(CStr(ZC_REBUILD_FILE_COUNT)) Then Call SetBlogHint(Empty,Empty,Empty)

	Dim strZC_REBUILD_FILE_INTERVAL
	strZC_REBUILD_FILE_INTERVAL=Request.Form("edtZC_REBUILD_FILE_INTERVAL")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_REBUILD_FILE_INTERVAL",strZC_REBUILD_FILE_INTERVAL)
	If UCase(strZC_REBUILD_FILE_INTERVAL)<>UCase(CStr(ZC_REBUILD_FILE_INTERVAL)) Then Call SetBlogHint(Empty,Empty,Empty)

	Dim strZC_UPLOAD_FILETYPE
	strZC_UPLOAD_FILETYPE=Request.Form("edtZC_UPLOAD_FILETYPE")
	Call SaveValueForSetting(strContent,True,"String","ZC_UPLOAD_FILETYPE",strZC_UPLOAD_FILETYPE)
	If UCase(strZC_UPLOAD_FILETYPE)<>UCase("""" & CStr(ZC_UPLOAD_FILETYPE) & """") Then Call SetBlogHint(Empty,Empty,Empty)

	Dim strZC_UPLOAD_FILESIZE
	strZC_UPLOAD_FILESIZE=Request.Form("edtZC_UPLOAD_FILESIZE")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_UPLOAD_FILESIZE",strZC_UPLOAD_FILESIZE)
	If UCase(strZC_UPLOAD_FILESIZE)<>UCase(CStr(ZC_UPLOAD_FILESIZE)) Then Call SetBlogHint(Empty,Empty,Empty)

	Dim strZC_COMMENT_VERIFY_ENABLE
	strZC_COMMENT_VERIFY_ENABLE=Request.Form("edtZC_COMMENT_VERIFY_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_COMMENT_VERIFY_ENABLE",strZC_COMMENT_VERIFY_ENABLE)
	If UCase(strZC_COMMENT_VERIFY_ENABLE)<>UCase(CStr(ZC_COMMENT_VERIFY_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)



	Dim strZC_MSG_COUNT
	strZC_MSG_COUNT=Request.Form("edtZC_MSG_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_MSG_COUNT",strZC_MSG_COUNT)
	If UCase(strZC_MSG_COUNT)<>UCase(CStr(ZC_MSG_COUNT)) Then Call SetBlogHint(Empty,True,Empty)

	Dim strZC_ARCHIVE_COUNT
	strZC_ARCHIVE_COUNT=Request.Form("edtZC_ARCHIVE_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_ARCHIVE_COUNT",strZC_ARCHIVE_COUNT)
	If UCase(strZC_ARCHIVE_COUNT)<>UCase(CStr(ZC_ARCHIVE_COUNT)) Then Call SetBlogHint(Empty,True,Empty)

	Dim strZC_PREVIOUS_COUNT
	strZC_PREVIOUS_COUNT=Request.Form("edtZC_PREVIOUS_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_PREVIOUS_COUNT",strZC_PREVIOUS_COUNT)
	If UCase(strZC_PREVIOUS_COUNT)<>UCase(CStr(ZC_PREVIOUS_COUNT)) Then Call SetBlogHint(Empty,True,Empty)

	Dim strZC_DISPLAY_COUNT
	strZC_DISPLAY_COUNT=Request.Form("edtZC_DISPLAY_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_DISPLAY_COUNT",strZC_DISPLAY_COUNT)
	If UCase(strZC_DISPLAY_COUNT)<>UCase(CStr(ZC_DISPLAY_COUNT)) Then Call SetBlogHint(Empty,True,Empty)

	Dim strZC_MANAGE_COUNT
	strZC_MANAGE_COUNT=Request.Form("edtZC_MANAGE_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_MANAGE_COUNT",strZC_MANAGE_COUNT)
	If UCase(strZC_MANAGE_COUNT)<>UCase(CStr(ZC_MANAGE_COUNT)) Then Call SetBlogHint(Empty,Empty,Empty)

	Dim strZC_RSS2_COUNT
	strZC_RSS2_COUNT=Request.Form("edtZC_RSS2_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_RSS2_COUNT",strZC_RSS2_COUNT)
	If UCase(strZC_RSS2_COUNT)<>UCase(CStr(ZC_RSS2_COUNT)) Then Call SetBlogHint(Empty,True,Empty)

	Dim strZC_SEARCH_COUNT
	strZC_SEARCH_COUNT=Request.Form("edtZC_SEARCH_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_SEARCH_COUNT",strZC_SEARCH_COUNT)
	If UCase(strZC_SEARCH_COUNT)<>UCase(CStr(ZC_SEARCH_COUNT)) Then Call SetBlogHint(Empty,Empty,Empty)

	Dim strZC_PAGEBAR_COUNT
	strZC_PAGEBAR_COUNT=Request.Form("edtZC_PAGEBAR_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_PAGEBAR_COUNT",strZC_PAGEBAR_COUNT)
	If UCase(strZC_PAGEBAR_COUNT)<>UCase(CStr(ZC_PAGEBAR_COUNT)) Then Call SetBlogHint(Empty,True,Empty)

	Dim strZC_USE_NAVIGATE_ARTICLE
	strZC_USE_NAVIGATE_ARTICLE=Request.Form("edtZC_USE_NAVIGATE_ARTICLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_USE_NAVIGATE_ARTICLE",strZC_USE_NAVIGATE_ARTICLE)
	If UCase(strZC_USE_NAVIGATE_ARTICLE)<>UCase(CStr(ZC_USE_NAVIGATE_ARTICLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_MUTUALITY_COUNT
	strZC_MUTUALITY_COUNT=Request.Form("edtZC_MUTUALITY_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_MUTUALITY_COUNT",strZC_MUTUALITY_COUNT)
	If UCase(strZC_MUTUALITY_COUNT)<>UCase(CStr(ZC_MUTUALITY_COUNT)) Then Call SetBlogHint(Empty,Empty,True)


	Dim strZC_UBB_LINK_ENABLE
	strZC_UBB_LINK_ENABLE=Request.Form("edtZC_UBB_LINK_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_UBB_LINK_ENABLE",strZC_UBB_LINK_ENABLE)
	If UCase(strZC_UBB_LINK_ENABLE)<>UCase(CStr(ZC_UBB_LINK_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_UBB_FONT_ENABLE
	strZC_UBB_FONT_ENABLE=Request.Form("edtZC_UBB_FONT_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_UBB_FONT_ENABLE",strZC_UBB_FONT_ENABLE)
	If UCase(strZC_UBB_FONT_ENABLE)<>UCase(CStr(ZC_UBB_FONT_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_UBB_CODE_ENABLE
	strZC_UBB_CODE_ENABLE=Request.Form("edtZC_UBB_CODE_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_UBB_CODE_ENABLE",strZC_UBB_CODE_ENABLE)
	If UCase(strZC_UBB_CODE_ENABLE)<>UCase(CStr(ZC_UBB_CODE_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_UBB_FACE_ENABLE
	strZC_UBB_FACE_ENABLE=Request.Form("edtZC_UBB_FACE_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_UBB_FACE_ENABLE",strZC_UBB_FACE_ENABLE)
	If UCase(strZC_UBB_FACE_ENABLE)<>UCase(CStr(ZC_UBB_FACE_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_UBB_IMAGE_ENABLE
	strZC_UBB_IMAGE_ENABLE=Request.Form("edtZC_UBB_IMAGE_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_UBB_IMAGE_ENABLE",strZC_UBB_IMAGE_ENABLE)
	If UCase(strZC_UBB_IMAGE_ENABLE)<>UCase(CStr(ZC_UBB_IMAGE_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_UBB_MEDIA_ENABLE
	strZC_UBB_MEDIA_ENABLE=Request.Form("edtZC_UBB_MEDIA_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_UBB_MEDIA_ENABLE",strZC_UBB_MEDIA_ENABLE)
	If UCase(strZC_UBB_MEDIA_ENABLE)<>UCase(CStr(ZC_UBB_MEDIA_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_UBB_FLASH_ENABLE
	strZC_UBB_FLASH_ENABLE=Request.Form("edtZC_UBB_FLASH_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_UBB_FLASH_ENABLE",strZC_UBB_FLASH_ENABLE)
	If UCase(strZC_UBB_FLASH_ENABLE)<>UCase(CStr(ZC_UBB_FLASH_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_UBB_TYPESET_ENABLE
	strZC_UBB_TYPESET_ENABLE=Request.Form("edtZC_UBB_TYPESET_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_UBB_TYPESET_ENABLE",strZC_UBB_TYPESET_ENABLE)
	If UCase(strZC_UBB_TYPESET_ENABLE)<>UCase(CStr(ZC_UBB_TYPESET_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_UBB_AUTOLINK_ENABLE
	strZC_UBB_AUTOLINK_ENABLE=Request.Form("edtZC_UBB_AUTOLINK_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_UBB_AUTOLINK_ENABLE",strZC_UBB_AUTOLINK_ENABLE)
	If UCase(strZC_UBB_AUTOLINK_ENABLE)<>UCase(CStr(ZC_UBB_AUTOLINK_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	'Dim strZC_AUTO_NEWLINE
	'strZC_AUTO_NEWLINE=Request.Form("edtZC_AUTO_NEWLINE")
	'Call SaveValueForSetting(strContent,True,"Boolean","ZC_AUTO_NEWLINE",strZC_AUTO_NEWLINE)
	'If UCase(strZC_AUTO_NEWLINE)<>UCase(CStr(ZC_AUTO_NEWLINE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_COMMENT_NOFOLLOW_ENABLE
	strZC_COMMENT_NOFOLLOW_ENABLE=Request.Form("edtZC_COMMENT_NOFOLLOW_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_COMMENT_NOFOLLOW_ENABLE",strZC_COMMENT_NOFOLLOW_ENABLE)
	If UCase(strZC_COMMENT_NOFOLLOW_ENABLE)<>UCase(CStr(ZC_COMMENT_NOFOLLOW_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_JAPAN_TO_HTML
	strZC_JAPAN_TO_HTML=Request.Form("edtZC_JAPAN_TO_HTML")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_JAPAN_TO_HTML",strZC_JAPAN_TO_HTML)
	If UCase(strZC_JAPAN_TO_HTML)<>UCase(CStr(ZC_JAPAN_TO_HTML)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_EMOTICONS_FILENAME
	strZC_EMOTICONS_FILENAME=Request.Form("edtZC_EMOTICONS_FILENAME")
	Call SaveValueForSetting(strContent,True,"String","ZC_EMOTICONS_FILENAME",strZC_EMOTICONS_FILENAME)
	If UCase(strZC_EMOTICONS_FILENAME)<>UCase("""" & CStr(ZC_EMOTICONS_FILENAME) & """") Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_EMOTICONS_FILESIZE
	strZC_EMOTICONS_FILESIZE=Request.Form("edtZC_EMOTICONS_FILESIZE")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_EMOTICONS_FILESIZE",strZC_EMOTICONS_FILESIZE)
	If UCase(strZC_EMOTICONS_FILESIZE)<>UCase(CStr(ZC_EMOTICONS_FILESIZE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_COMMENT_REVERSE_ORDER_EXPORT
	strZC_COMMENT_REVERSE_ORDER_EXPORT=Request.Form("edtZC_COMMENT_REVERSE_ORDER_EXPORT")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_COMMENT_REVERSE_ORDER_EXPORT",strZC_COMMENT_REVERSE_ORDER_EXPORT)
	If UCase(strZC_COMMENT_REVERSE_ORDER_EXPORT)<>UCase(CStr(ZC_COMMENT_REVERSE_ORDER_EXPORT)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_GUESTBOOK_CONTENT
	strZC_GUESTBOOK_CONTENT=Replace(Replace(Request.Form("edtZC_GUESTBOOK_CONTENT"),vbCr,""),vbLf,"")
	Call SaveValueForSetting(strContent,True,"String","ZC_GUESTBOOK_CONTENT",strZC_GUESTBOOK_CONTENT)
	If UCase(strZC_GUESTBOOK_CONTENT)<>UCase("""" & CStr(ZC_GUESTBOOK_CONTENT) & """") Then Call SetBlogHint(Empty,Empty,Empty)

	Dim strZC_CUSTOM_DIRECTORY_ENABLE
	strZC_CUSTOM_DIRECTORY_ENABLE=Request.Form("edtZC_CUSTOM_DIRECTORY_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_CUSTOM_DIRECTORY_ENABLE",strZC_CUSTOM_DIRECTORY_ENABLE)
	If UCase(strZC_CUSTOM_DIRECTORY_ENABLE)<>UCase(CStr(ZC_CUSTOM_DIRECTORY_ENABLE)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_CUSTOM_DIRECTORY_ANONYMOUS
	strZC_CUSTOM_DIRECTORY_ANONYMOUS=Request.Form("edtZC_CUSTOM_DIRECTORY_ANONYMOUS")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_CUSTOM_DIRECTORY_ANONYMOUS",strZC_CUSTOM_DIRECTORY_ANONYMOUS)
	If UCase(strZC_CUSTOM_DIRECTORY_ANONYMOUS)<>UCase(CStr(ZC_CUSTOM_DIRECTORY_ANONYMOUS)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_CUSTOM_DIRECTORY_REGEX
	strZC_CUSTOM_DIRECTORY_REGEX=Request.Form("edtZC_CUSTOM_DIRECTORY_REGEX")
	If strZC_CUSTOM_DIRECTORY_ANONYMOUS="True" Then
		If InStr(strZC_CUSTOM_DIRECTORY_REGEX,"{%id%}")=0 And InStr(strZC_CUSTOM_DIRECTORY_REGEX,"{%alias%}")=0 Then
			strZC_CUSTOM_DIRECTORY_REGEX=strZC_CUSTOM_DIRECTORY_REGEX & "{%id%}"
		End If
	End If
	Call SaveValueForSetting(strContent,True,"String","ZC_CUSTOM_DIRECTORY_REGEX",strZC_CUSTOM_DIRECTORY_REGEX)
	If UCase(strZC_CUSTOM_DIRECTORY_REGEX)<>UCase("""" & CStr(ZC_CUSTOM_DIRECTORY_REGEX) & """") Then Call SetBlogHint(Empty,Empty,True)


	'Dim strZC_IE_DISPLAY_WAP
	'strZC_IE_DISPLAY_WAP=Request.Form("edtZC_IE_DISPLAY_WAP")
	'Call SaveValueForSetting(strContent,True,"Boolean","ZC_IE_DISPLAY_WAP",strZC_IE_DISPLAY_WAP)

	Dim strZC_DISPLAY_COUNT_WAP
	strZC_DISPLAY_COUNT_WAP=Request.Form("edtZC_DISPLAY_COUNT_WAP")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_DISPLAY_COUNT_WAP",strZC_DISPLAY_COUNT_WAP)

	Dim strZC_COMMENT_COUNT_WAP
	strZC_COMMENT_COUNT_WAP=Request.Form("edtZC_COMMENT_COUNT_WAP")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_COMMENT_COUNT_WAP",strZC_COMMENT_COUNT_WAP)

	Dim strZC_PAGEBAR_COUNT_WAP
	strZC_PAGEBAR_COUNT_WAP=Request.Form("edtZC_PAGEBAR_COUNT_WAP")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_PAGEBAR_COUNT_WAP",strZC_PAGEBAR_COUNT_WAP)

	Dim strZC_SINGLE_SIZE_WAP
	strZC_SINGLE_SIZE_WAP=Request.Form("edtZC_SINGLE_SIZE_WAP")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_SINGLE_SIZE_WAP",strZC_SINGLE_SIZE_WAP)

	Dim strZC_SINGLE_PAGEBAR_COUNT_WAP
	strZC_SINGLE_PAGEBAR_COUNT_WAP=Request.Form("edtZC_SINGLE_PAGEBAR_COUNT_WAP")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_SINGLE_PAGEBAR_COUNT_WAP",strZC_SINGLE_PAGEBAR_COUNT_WAP)

	Dim strZC_COMMENT_PAGEBAR_COUNT_WAP
	strZC_COMMENT_PAGEBAR_COUNT_WAP=Request.Form("edtZC_COMMENT_PAGEBAR_COUNT_WAP")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_COMMENT_PAGEBAR_COUNT_WAP",strZC_COMMENT_PAGEBAR_COUNT_WAP)

	Dim strZC_FILENAME_WAP
	strZC_FILENAME_WAP=Request.Form("edtZC_FILENAME_WAP")
	Call SaveValueForSetting(strContent,True,"String","ZC_FILENAME_WAP",strZC_FILENAME_WAP)

	Dim strZC_WAPCOMMENT_ENABLE
	strZC_WAPCOMMENT_ENABLE=Request.Form("edtZC_WAPCOMMENT_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_WAPCOMMENT_ENABLE",strZC_WAPCOMMENT_ENABLE)

	Dim strZC_UPLOAD_DIRBYMONTH
	strZC_UPLOAD_DIRBYMONTH=Request.Form("edtZC_UPLOAD_DIRBYMONTH")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_UPLOAD_DIRBYMONTH",strZC_UPLOAD_DIRBYMONTH)
	If UCase(ZC_UPLOAD_DIRBYMONTH)<>UCase(CStr(ZC_UPLOAD_DIRBYMONTH)) Then Call SetBlogHint(Empty,Empty,True)


	Dim strZC_IMAGE_WIDTH
	strZC_IMAGE_WIDTH=Request.Form("edtZC_IMAGE_WIDTH")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_IMAGE_WIDTH",strZC_IMAGE_WIDTH)
	If UCase(strZC_IMAGE_WIDTH)<>UCase(CStr(ZC_IMAGE_WIDTH)) Then Call SetBlogHint(Empty,Empty,True)

	Dim strZC_RSS_EXPORT_WHOLE
	strZC_RSS_EXPORT_WHOLE=Request.Form("edtZC_RSS_EXPORT_WHOLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_RSS_EXPORT_WHOLE",strZC_RSS_EXPORT_WHOLE)
	If UCase(strZC_RSS_EXPORT_WHOLE)<>UCase(CStr(ZC_RSS_EXPORT_WHOLE)) Then Call SetBlogHint(Empty,True,Empty)

	Dim strZC_COMMENT_TURNOFF
	strZC_COMMENT_TURNOFF=Request.Form("edtZC_COMMENT_TURNOFF")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_COMMENT_TURNOFF",strZC_COMMENT_TURNOFF)

	Dim strZC_TRACKBACK_TURNOFF
	strZC_TRACKBACK_TURNOFF=Request.Form("edtZC_TRACKBACK_TURNOFF")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_TRACKBACK_TURNOFF",strZC_TRACKBACK_TURNOFF)

	Dim strZC_GUEST_REVERT_COMMENT_ENABLE
	strZC_GUEST_REVERT_COMMENT_ENABLE=Request.Form("edtZC_GUEST_REVERT_COMMENT_ENABLE")
	Call SaveValueForSetting(strContent,True,"Boolean","ZC_GUEST_REVERT_COMMENT_ENABLE",strZC_GUEST_REVERT_COMMENT_ENABLE)


	Dim strZC_VERIFYCODE_WIDTH
	strZC_VERIFYCODE_WIDTH=Request.Form("edtZC_VERIFYCODE_WIDTH")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_VERIFYCODE_WIDTH",strZC_VERIFYCODE_WIDTH)

	Dim strZC_VERIFYCODE_HEIGHT
	strZC_VERIFYCODE_HEIGHT=Request.Form("edtZC_VERIFYCODE_HEIGHT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_VERIFYCODE_HEIGHT",strZC_VERIFYCODE_HEIGHT)

	Dim strZC_VERIFYCODE_STRING
	strZC_VERIFYCODE_STRING=Request.Form("edtZC_VERIFYCODE_STRING")
	Call SaveValueForSetting(strContent,True,"String","ZC_VERIFYCODE_STRING",strZC_VERIFYCODE_STRING)
	If UCase(strZC_VERIFYCODE_STRING)<>UCase(CStr(ZC_VERIFYCODE_STRING)) Then Application.Lock : Application(ZC_BLOG_CLSID & "VERIFY_NUMBER")=Empty : Application.UnLock

	Dim strZC_RECENT_COMMENT_WORD_MAX
	strZC_RECENT_COMMENT_WORD_MAX=Request.Form("edtZC_RECENT_COMMENT_WORD_MAX")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_RECENT_COMMENT_WORD_MAX",strZC_RECENT_COMMENT_WORD_MAX)
	If UCase(strZC_RECENT_COMMENT_WORD_MAX)<>UCase(CStr(ZC_RECENT_COMMENT_WORD_MAX)) Then Call SetBlogHint(Empty,True,Empty)

	Dim strZC_TAGS_DISPLAY_COUNT
	strZC_TAGS_DISPLAY_COUNT=Request.Form("edtZC_TAGS_DISPLAY_COUNT")
	Call SaveValueForSetting(strContent,True,"Numeric","ZC_TAGS_DISPLAY_COUNT",strZC_TAGS_DISPLAY_COUNT)
	If UCase(strZC_TAGS_DISPLAY_COUNT)<>UCase(CStr(ZC_TAGS_DISPLAY_COUNT)) Then Call SetBlogHint(Empty,True,Empty)

'	Dim str<#>
'	str<#>=Request.Form("edt<#>")
'	Call SaveValueForSetting(strContent,True,"Boolean","<#>",str<#>)

	Call SaveToFile(BlogPath & "zb_users/c_option.asp",strContent,"utf-8",False)

	'Call MakeBlogReBuild_Core()

	SaveSetting=True

	Err.Clear

End Function
'*********************************************************




'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    Post Tag
'*********************************************************
Function PostTag()

	Dim objTag
	Set objTag=New TTag
	objTag.ID=Request.Form("edtID")
	objTag.Name=Request.Form("edtName")
	objTag.Order=Request.Form("edtOrder")
	objTag.Intro=Request.Form("edtIntro")
	objTag.Alias=Request.Form("edtAlias")
	objTag.TemplateName=Request.Form("edtTemplate")

	'接口
	Call Filter_Plugin_PostTag_Core(objTag)

	If objTag.Post Then
		Call ScanTagCount("{"&objTag.ID&"}")
		PostTag=True
		Call Filter_Plugin_PostTag_Succeed(objTag)
	End If
	Set objTag=Nothing

End Function
'*********************************************************




'*********************************************************
' 目的：    Del Tag
'*********************************************************
Function DelTag(intID)

	Dim objTag
	Set objTag=New TTag
	objTag.ID=intID
	If objTag.Del Then DelTag=True
	Set objTag=Nothing

End Function
'*********************************************************




'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    Post KeyWord
'*********************************************************
Function PostKeyWord()

End Function
'*********************************************************




'*********************************************************
' 目的：    Del Tag
'*********************************************************
Function DelKeyWord(intID)

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function PostSiteFile(tpath)
	If Instr(lcase(tpath),"global.asa") then postsitefile=false:exit function
	Dim txaContent
	txaContent=Request.Form("txaContent")

	If IsEmpty(txaContent) Then txaContent=Null

	If Not IsNull(tpath) Then

		If Not IsNull(txaContent) Then

			Call SaveToFile(BlogPath & tpath,txaContent,"utf-8",False)

			PostSiteFile=True

		End IF

	End If


End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function DelSiteFile(tpath)

	Dim Fso
	Set Fso = Createobject("Scripting.Filesystemobject")
	If Fso.FileExists(BlogPath & tpath) Then
		Fso.Deletefile(BlogPath & tpath)
		Set Fso = Nothing

		DelSiteFile=True

		Exit Function
	Else
		Set Fso = Nothing
		Exit Function
	End If

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function GetRealUrlofTrackBackUrl(intID)

	Call CheckParameter(intID,"int",0)

	If IsEmpty(Request.Form("edtCheckOut")) Then

		Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><link rel=""stylesheet"" rev=""stylesheet"" href=""CSS/admin.css"" type=""text/css"" media=""screen"" /></head><body>"

		Response.Write "<div id=""divMain""><div class=""Header"">" & ZC_MSG145 & "</div>"
		Response.Write "<div id=""divMain2""><form method=""post""  name=""edit"" id=""edit"" action="""& ZC_BLOG_HOST & "cmd.asp?act=gettburl&id=" & intID &""">"
		Response.Write "<p></p>"
		Response.Write "<p>"& ZC_MSG161 &"</p>"
		Response.Write "<p><img style=""border:1px solid black""  src=""function/c_validcode.asp?name=gettburlvalid"" height="""&ZC_VERIFYCODE_HEIGHT&""" width="""&ZC_VERIFYCODE_WIDTH&""" alt="""" title=""""/>&nbsp;<input type=""text"" id=""edtCheckOut"" name=""edtCheckOut"" size=""30"" />&nbsp;<input class=""button"" type=""submit"" value=""" & ZC_MSG087 & """ id=""btnPost""></p>"
		Response.Write "<p></p>"
		Response.Write "</form></div></div>"
		Response.Write "</body></html>"

	ElseIf CheckVerifyNumber(Request.Form("edtCheckOut"))=True Then


		Response.Write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><link rel=""stylesheet"" rev=""stylesheet"" href=""CSS/admin.css"" type=""text/css"" media=""screen"" /></head><body>"

		Response.Write "<div id=""divMain""><div class=""Header"">" & ZC_MSG145 & "</div>"
		Response.Write "<div id=""divMain2""><form method=""post""  name=""edit"" id=""edit"" action="""& ZC_BLOG_HOST & "cmd.asp?act=gettburl&id=" & intID &""">"
		Response.Write "<p></p>"
		Response.Write "<p>" & ZC_BLOG_HOST & "cmd.asp?act=tb&amp;id=" & intID & "&amp;key=" & GetVerifyNumber() &"</p>"
		Response.Write "<p></p>"
		Response.Write "</form></div></div>"
		Response.Write "</body></html>"

		Session("gettburlvalid")=Empty

	Else

		Call ShowError(38)

	End If

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function DelCommentBatch()
	On Error Resume Next
	Dim i,j
	Dim s,t
	Dim aryArticle()
	s=Request.Form("edtBatch")
	t=Split(s,",")

	ReDim Preserve aryArticle(UBound(t))
	For j=0 To UBound(t)-1
		aryArticle(j)=0
	Next

	Dim objComment
	Dim objArticle
	Dim objRs

	For i=0 To UBound(t)-1
		Set objComment=New TComment
		If objComment.LoadInfobyID(t(i)) Then
			If objComment.log_ID>0 Then
				Dim objTestArticle
				Set objTestArticle=New TArticle
				If objTestArticle.LoadInfobyID(objComment.log_ID) Then

					For j=0 To UBound(t)-1
						If aryArticle(j)=0 Then
							aryArticle(j)=objComment.log_ID
						End If
						If aryArticle(j)=objComment.log_ID Then Exit For
					Next

					If Not((objComment.AuthorID=BlogUser.ID) Or (objTestArticle.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Exit Function
				Else
					Call ShowError(9)
				End If
				Set objTestArticle=Nothing
			Else
				If Not((objComment.log_ID=0) And (CheckRights("GuestBookMng")=True)) Then Exit Function
			End If

			DelChild objComment.ID
			objComment.Del

		End If
		Set objComment=Nothing
	Next


	For j=0 To UBound(t)-1
		If aryArticle(j)>0 Then
			Call BuildArticle(aryArticle(j),False,False)
		End If
	Next

	BlogReBuild_Comments
	BlogReBuild_GuestComments
	DelCommentBatch=True

End Function
'*********************************************************
'*********************************************************
' 目的：
'*********************************************************
Function DelChild(ID)
	Dim objRs,tmpParentID,tmpID
	Set objRs=objConn.Execute("SELECT comm_ParentID,comm_ID FROM [blog_Comment] WHERE comm_ParentID="&ID)
	Do Until objRS.bof Or objRS.eof
		tmpParentID=clng(objRs("comm_ParentID"))
		tmpID=clng(objRs("comm_ID"))

		If tmpParentID>0 Then
			DelChild(tmpID)
		End If
		objConn.Execute "DELETE FROM [blog_Comment] WHERE comm_ID="&tmpID
		objRS.MoveNext
	Loop
End Function
'*********************************************************

'*********************************************************
' 目的：
'*********************************************************
Function DelTrackBackBatch()

	Dim i,j
	Dim s,t
	Dim aryArticle()
	s=Request.Form("edtBatch")
	t=Split(s,",")

	ReDim Preserve aryArticle(UBound(t))
	For j=0 To UBound(t)-1
		aryArticle(j)=0
	Next

	Dim objTrackBack
	Dim objArticle

	Set objArticle=New TArticle


	For i=0 To UBound(t)-1
		Set objTrackBack=New TTrackBack
		If objTrackBack.LoadInfobyID(t(i)) Then
			Dim objTestArticle
			Set objTestArticle=New TArticle
			If objTestArticle.LoadInfobyID(objTrackBack.log_ID) Then

				For j=0 To UBound(t)-1
					If aryArticle(j)=0 Then
						aryArticle(j)=objTrackBack.log_ID
					End If
					If aryArticle(j)=objTrackBack.log_ID Then Exit For
				Next

				If Not((objTestArticle.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Exit Function
			Else
				Call ShowError(9)
			End If
			Set objTestArticle=Nothing
			objTrackBack.Del
		End If
		Set objTrackBack=Nothing
	Next


	For j=0 To UBound(t)-1
		If aryArticle(j)>0 Then
			Call BuildArticle(aryArticle(j),False,False)
		End If
	Next

	BlogReBuild_TrackBacks
	DelTrackBackBatch=True

End Function
'*********************************************************




'*********************************************************
' 目的：     文件删除
'*********************************************************
Function DelFileBatch()

	Dim i,j
	Dim s,t

	s=Request.Form("edtBatch")
	t=Split(s,",")

	Dim objUpLoadFile

	For i=0 To UBound(t)-1
		t(i)=CLng(t(i))
		If t(i)>0 Then
			Set objUpLoadFile=New TUpLoadFile
			If objUpLoadFile.LoadInfoByID(t(i)) Then
				If (objUpLoadFile.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True) Then
					objUpLoadFile.Del
				End If
			Else
				Exit Function
			End If
			Set objUpLoadFile=Nothing
		End If
	Next

	DelFileBatch=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Save Theme Setting
'*********************************************************
Function SaveTheme()

	Dim i,j
	Dim s,t
	Dim strContent

	strContent=LoadFromFile(BlogPath & "zb_users/c_custom.asp","utf-8")

	Dim strZC_BLOG_CSS
	Dim strZC_BLOG_THEME

	strZC_BLOG_CSS=Request.Form("edtZC_BLOG_CSS")
	strZC_BLOG_THEME=Request.Form("edtZC_BLOG_THEME")

	Call ScanPluginToThemeFile(strZC_BLOG_CSS,strZC_BLOG_THEME)

	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_CSS",strZC_BLOG_CSS)
	Call SaveValueForSetting(strContent,True,"String","ZC_BLOG_THEME",strZC_BLOG_THEME)

	If UCase(strZC_BLOG_CSS)<>UCase("""" & CStr(ZC_BLOG_CSS) & """") Then Call SetBlogHint(Empty,True,Empty)
	If UCase(strZC_BLOG_THEME)<>UCase("""" & CStr(ZC_BLOG_THEME) & """") Then Call SetBlogHint(Empty,True,True):Call UninstallPlugin(ZC_BLOG_THEME)

	Call SaveToFile(BlogPath & "zb_users/c_custom.asp",strContent,"utf-8",False)

	Call MakeBlogReBuild_Core()

	SaveTheme=True

End Function
'*********************************************************




'*********************************************************
' 目的：    Save Links
'*********************************************************
Function SaveLink()

	Dim tpath
	Dim txaContent
	Dim strContent


	tpath="./ZB_USERS/INCLUDE/link.asp"
	txaContent=Request.Form("txaContent_Link")

	If IsEmpty(txaContent) Then txaContent=Null
	If Not IsNull(tpath) Then
		If Not IsNull(txaContent) Then
			Call SaveToFile(BlogPath & tpath,txaContent,"utf-8",False)
		End IF
	End If

	tpath="./ZB_USERS/INCLUDE/favorite.asp"
	txaContent=Request.Form("txaContent_Favorite")

	If IsEmpty(txaContent) Then txaContent=Null
	If Not IsNull(tpath) Then
		If Not IsNull(txaContent) Then
			Call SaveToFile(BlogPath & tpath,txaContent,"utf-8",False)
		End IF
	End If

	tpath="./ZB_USERS/INCLUDE/misc.asp"
	txaContent=Request.Form("txaContent_Misc")

	If IsEmpty(txaContent) Then txaContent=Null
	If Not IsNull(tpath) Then
		If Not IsNull(txaContent) Then
			Call SaveToFile(BlogPath & tpath,txaContent,"utf-8",False)
		End IF
	End If

	tpath="./ZB_USERS/INCLUDE/navbar.asp"
	txaContent=Request.Form("txaContent_Navbar")

	If IsEmpty(txaContent) Then txaContent=Null
	If Not IsNull(tpath) Then
		If Not IsNull(txaContent) Then

			strContent=LoadFromFile(BlogPath & tpath,"utf-8")
			If txaContent<>strContent Then
				Call SetBlogHint(Empty,True,True)
			End If

			Call SaveToFile(BlogPath & tpath,txaContent,"utf-8",False)

		End IF
	End If

	Call SetBlogHint(Empty,True,Empty)

	Call MakeBlogReBuild_Core()

	SaveLink=True

End Function
'*********************************************************




'*********************************************************
' 目的：    ActivePlugIn By Name
'*********************************************************
Function ActivePlugInByName(strPluginName)

	Dim s,i,t,b
	s= ZC_USING_PLUGIN_LIST

	If s="" Then
		s=strPluginName
	Else
		t=Split(ZC_USING_PLUGIN_LIST,"|")
		For i=LBound(t) To UBound(t)
			If UCase(t(i))=UCase(strPluginName) Then
				b=True
			End If
		Next
		If b<>True Then
			s=s & "|" & strPluginName
		End If
	End If


	Dim strContent
	Dim strZC_USING_PLUGIN_LIST

	strContent=LoadFromFile(BlogPath & "zb_users/c_option.asp","utf-8")

	strZC_USING_PLUGIN_LIST=s

	Call SaveValueForSetting(strContent,True,"String","ZC_USING_PLUGIN_LIST",strZC_USING_PLUGIN_LIST)

	Call SaveToFile(BlogPath & "zb_users/c_option.asp",strContent,"utf-8",False)

	Call ScanPluginToIncludeFile(s)

	ActivePlugInByName=True

End Function
'*********************************************************




'*********************************************************
' 目的：    DisablePlugIn By Name
'*********************************************************
Function DisablePlugInByName(strPluginName)

	Call UninstallPlugin(strPluginName)

	Dim s,i,t

	s=Split(ZC_USING_PLUGIN_LIST,"|")

	For i=LBound(s) To UBound(s)

		If UCase(s(i))<>UCase(strPluginName) Then

			If t="" Then
				t=s(i)
			Else
				t=t & "|" & s(i)
			End If

		End If

	Next


	Dim strContent
	Dim strZC_USING_PLUGIN_LIST

	strContent=LoadFromFile(BlogPath & "zb_users/c_option.asp","utf-8")

	strZC_USING_PLUGIN_LIST=t

	Call SaveValueForSetting(strContent,True,"String","ZC_USING_PLUGIN_LIST",strZC_USING_PLUGIN_LIST)

	Call SaveToFile(BlogPath & "zb_users/c_option.asp",strContent,"utf-8",False)

	Call ScanPluginToIncludeFile(t)

	DisablePlugInByName=True

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function ScanPluginToIncludeFile(newZC_USING_PLUGIN_LIST)

	On Error Resume Next

	Dim aryPL,i,j,s,t
	aryPL=Split(newZC_USING_PLUGIN_LIST,"|")

	If newZC_USING_PLUGIN_LIST<>"" Then
		i=UBound(aryPL)
	Else
		i=0
	End If


	Dim objXmlFile,strXmlFile
	Dim fso, f, f1, fc
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "zb_users/plugin/")
	Set fc = f.SubFolders
	For j=0 To i
		If fso.FileExists(BlogPath & "zb_users/plugin/" & aryPL(j) & "/" & "plugin.xml") Then

			strXmlFile =BlogPath & "zb_users/plugin/" & aryPL(j) & "/" & "plugin.xml"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else
					If CheckPluginStateByNewValue(objXmlFile.documentElement.selectSingleNode("id").text,newZC_USING_PLUGIN_LIST) Then
						If Trim(objXmlFile.documentElement.selectSingleNode("include").text)<>"" Then
							If (fso.FileExists(BlogPath & "zb_users/plugin/" & objXmlFile.documentElement.selectSingleNode("id").text & "/" & objXmlFile.documentElement.selectSingleNode("include").text)) Then
								t="<!-- #include file="""& objXmlFile.documentElement.selectSingleNode("id").text &"/"& objXmlFile.documentElement.selectSingleNode("include").text &""" -->"
								If InStr(s,t)=0 Then
									s=s & t  & vbCrLf
								End If
							End If
						End If
					End If
				End If
			End If
			Set objXmlFile=Nothing
		End If
	Next

	Call SaveToFile(BlogPath & "zb_users/PLUGIN/p_include.asp",s,"utf-8",False)

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function ScanPluginToThemeFile(newZC_BLOG_CSS,newZC_BLOG_THEME)

	On Error Resume Next

	Dim objXmlFile,strXmlFile,s

	strXmlFile =BlogPath & "zb_users/theme/" & newZC_BLOG_THEME & "/" & "theme.xml"

	Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
	objXmlFile.async = False
	objXmlFile.ValidateOnParse=False
	objXmlFile.load(strXmlFile)
	If objXmlFile.readyState=4 Then
		If objXmlFile.parseError.errorCode <> 0 Then
		Else
			If LCase(objXmlFile.documentElement.selectSingleNode("id").text)=LCase(newZC_BLOG_THEME) Then
				Dim fso
				Set fso = CreateObject("Scripting.FileSystemObject")
				If (fso.FileExists(BlogPath & "zb_users/theme/" & objXmlFile.documentElement.selectSingleNode("id").text &"/plugin/" & objXmlFile.documentElement.selectSingleNode("plugin/include").text)) Then
					If Trim(objXmlFile.documentElement.selectSingleNode("plugin/include").text)<>"" Then
						s=s & "<!-- #include file=""../theme/"& objXmlFile.documentElement.selectSingleNode("id").text &"/plugin/"& objXmlFile.documentElement.selectSingleNode("plugin/include").text &""" -->" & vbCrLf
					End If
				End If
			End If
		End If
	End If
	Set objXmlFile=Nothing

	Call SaveToFile(BlogPath & "zb_users/PLUGIN/p_theme.asp",s,"utf-8",False)

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：
'*********************************************************
Function ScanTagCount(strTags)

	On Error Resume Next

	Dim t,i,s
	Dim objRS,j,k

	If strTags<>"" Then
		s=strTags
		s=Replace(s,"}","")
		t=Split(s,"{")

		For i=LBound(t) To UBound(t)
			If t(i)<>"" Then
				k=Tags(t(i)).ID

				Set objRS=objConn.Execute("SELECT COUNT([log_ID]) FROM [blog_Article] WHERE [log_Level]>1 AND [log_Tag] LIKE '%{" & k & "}%'")
				j=objRS(0)
				objConn.Execute("UPDATE [blog_Tag] SET [tag_Count]="&j&" WHERE [tag_ID] =" & k)
				Set objRS=Nothing

			End If
		Next

		s=Join(t,",")
		s=Right(s,Len(s)-1)

		strTags=s
	End If

	Err.Clear

End Function
'*********************************************************

%>