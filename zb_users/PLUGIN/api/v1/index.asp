<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../../c_option.asp" -->
<!-- #include file="../../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../p_config.asp" -->
<!-- #include file="../JSON.asp" -->
<!-- #include file="function.asp" -->
<%
Response.ContentType="application/json"
Call System_Initialize()

Dim objConfig,data_export,errcode,body_array,request_array(50),msg,ret:errcode="0":msg="true":ret=0
Set objConfig=New TConfig:Set data_export = jsObject():Set body_array = jsObject()
objConfig.Load("api")

If CheckPluginState("api")=False Then Call ShowApiError():Response.End

Dim strApi
strApi = Request("api")

'api检查
'If (strApi<>"tb") And (strApi<>"search") Then Call CheckReference("")

Select Case strApi
	'命令列表
	Case "verify"
		Call api_Verify()

	Case "blog_info"
		Call api_blog_info()

	Case "blog_set"
		Call api_blog_set()

	Case "cate_add"
		Call api_cate_add()

	Case "cate_del"
		Call api_cate_del()

	Case "cate_edit"
		Call api_cate_edit()

	Case "cate_list"
		Call api_cate_list()

	Case "comment_add"
		Call api_comment_add()

	Case "comment_del"
		Call api_comment_del()

	Case "comment_edit"
		Call api_comment_edit()

	Case "comment_get"
		Call api_comment_get()

	Case "comment_list"
		Call api_comment_list()

	Case "file_add"
		Call api_file_add()

	Case "file_del"
		Call api_file_del()

	Case "file_list"
		Call api_file_list()

	Case "page_list"
		Call api_page_list()

	Case "post_add"
		Call api_post_add()

	Case "post_del"
		Call api_post_del()

	Case "post_edit"
		Call api_post_edit()

	Case "post_get"
		Call api_post_get()

	Case "post_list"
		Call api_post_list()

	Case "sidebar_list"
		Call api_sidebar_list()
		
	Case "sidebar_get"
		Call api_sidebar_get()

	Case "tag_add"
		Call api_tag_add()

	Case "tag_del"
		Call api_tag_del()

	Case "tag_edit"
		Call api_tag_edit()

	Case "tag_list"
		Call api_tag_list()

	Case "user_edit"
		Call api_user_edit()

	Case "user_list"
		Call api_user_list()
End Select



'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: api插件未启用
' 参数: 
'*********************************************************
Function ShowApiError
	errcode="999"
	msg="api plugin is unavailable"
	ret=1

	Call Public_date()

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 验证key和连接
' 参数: ?keyid=0839609affdbb82c9884fc05a2dfcd18&keysecretmd5=a9b0772d4900afd67178b14579b008ca&post=111
'*********************************************************
Function api_Verify
	Verify()
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取博客信息
' 参数: 
'*********************************************************
Function api_blog_info
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	body_array("blog_title")=ZC_BLOG_TITLE
	body_array("blog_subtitle")=ZC_BLOG_SUBTITLE
	body_array("blog_url")=BlogHost
	body_array("blog_master")=ZC_BLOG_MASTER
	body_array("blog_language")=ZC_BLOG_LANGUAGE
	body_array("blog_languagepack")=ZC_BLOG_LANGUAGEPACK
	body_array("blog_version")=BlogVersion
	body_array("blog_theme")=ZC_BLOG_THEME
	body_array("blog_copyright")=ZC_BLOG_COPYRIGHT

	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取blog设置
' 参数: 
'*********************************************************
Function api_blog_set
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function

	GetBlogSet()
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 添加分类
' 参数: 
'*********************************************************
Function api_cate_add
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 删除分类
' 参数: 
'*********************************************************
Function api_cate_del
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	If Request.Form("id")<>""Then
		If Request.Form("id")=0 Then
			errcode="111":msg="cate 0 could't not be del."
		Else
			Dim objCategory
			Set objCategory=New TCategory

			If objCategory.LoadInfobyID(Request.Form("id")) Then
				If objCategory.Del Then 
					body_array("del_type")=0
				Else
					errcode="111":msg="cate del wrong."
				End If 
			End If
			Set objCategory=Nothing
		End If
	Else
		errcode="111":msg="post id is empty."
	End If
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 编辑分类
' 参数: 
'*********************************************************
Function api_cate_edit
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取分类列表
' 参数: 
'*********************************************************
Function api_cate_list
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 添加评论
' 参数: 
'*********************************************************
Function api_comment_add
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 删除评论
' 参数: 
'*********************************************************
Function api_comment_del
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 修改评论内容
' 参数: 
'*********************************************************
Function api_comment_edit
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取指定评论内容
' 参数: 
'*********************************************************
Function api_comment_get
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取评论列表
' 参数: 
'*********************************************************
Function api_comment_list
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 上传附件
' 参数: 
'*********************************************************
Function api_file_add
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 删除附件
' 参数: 
'*********************************************************
Function api_file_del
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	request_array(0)=Request.Form("fileid")
	
	Dim objUpLoadFile
	Set objUpLoadFile=New TUpLoadFile

	If objUpLoadFile.LoadInfoByID(request_array(0)) Then
		If objUpLoadFile.Del Then body_array("del_type")=0
	Else
		errcode="002":msg="file is not exist."
	End If
	Set objUpLoadFile=Nothing
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取附件列表
' 参数: 
'*********************************************************
Function api_file_list
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	Dim objRS,info_array,Doi:Doi=0
	Set objRS=objConn.Execute("SELECT * FROM [blog_UpLoad] ORDER BY [ul_PostTime] DESC")
	Do Until objRS.Eof
		Set info_array = jsObject()
		info_array("ID")=objRS("ul_Id")
		info_array("AuthorID")=objRS("ul_AuthorID")
		info_array("FileSize")=objRS("ul_FileSize")
		info_array("FileName")=objRS("ul_FileName")
		info_array("PostTime")=objRS("ul_PostTime")
		info_array("Meta")=objRS("ul_Meta")
		Set body_array(Doi) = info_array
		Doi=Doi+1
		objRS.MoveNext
	Loop
	
	Set objRS=Nothing
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取独立页面列表
' 参数: 
'*********************************************************
Function api_page_list
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	Dim objRS,info_array,Doi:Doi=0
	Set objRS=objConn.Execute("SELECT * FROM [blog_Article] WHERE ([log_Type]=1) AND ([log_Level]>0) AND (1=1)  ORDER BY [log_PostTime] DESC")
	Do Until objRS.Eof
		Set info_array = jsObject()
		info_array("ID")=objRS("log_ID")
		info_array("AuthorID")=objRS("log_AuthorID")
		info_array("Title")=objRS("log_Title")
		info_array("PostTime")=objRS("log_PostTime")
		info_array("CommNums")=objRS("log_CommNums")
		info_array("Level")=objRS("log_Level")
		Set body_array(Doi) = info_array
		Doi=Doi+1
		objRS.MoveNext
	Loop
	
	Set objRS=Nothing

	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 添加文章或独立页面
' 参数: 
'*********************************************************
Function api_post_add
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 删除指定ID文章或独立页面
' 参数: 
'*********************************************************
Function api_post_del
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	Dim strTag
	If Request.Form("id")<>""Then
		Dim objTestArticle
		Set objTestArticle=New TArticle
		If objTestArticle.LoadInfobyID(Request.Form("id")) Then
			strTag=objTestArticle.Tag
		Else
			errcode="111":msg="post id is wrong."
		End If
		Set objTestArticle=Nothing
	Else
		errcode="111":msg="post id is empty."
	End If
	
	Dim objArticle
	Set objArticle=New TArticle
	If objArticle.LoadInfoByID(Request.Form("ID")) Then
		If objArticle.Del Then 
			body_array("del_type")=0
			objArticle.Statistic
		Else
			errcode="111":msg="post del wrong."
		End If 
		'Call ScanTagCount(strTag)
		Call BlogReBuild_Comments
		Call BlogReBuild_Default
	End If
	Set objArticle=Nothing

	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 修改指定ID文章或独立页面
' 参数: 
'*********************************************************
Function api_post_edit
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	
	




End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取制定ID文章或独立页面
' 参数: 
'*********************************************************
Function api_post_get
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	Dim objArticle
	Set objArticle=New TArticle
	
	request_array(0)=Request.Form("id")
	If request_array(0)<>"" Then
		request_array(0)=CInt(request_array(0))
		If request_array(0)=0 Then errcode="111":msg="post id is wrong."
	Else
		errcode="111":msg="post id is wrong."
	End If
	If ErrorCheck() Then Exit Function
	
	If objArticle.LoadInfoByID(request_array(0)) Then
		body_array("postid")=objArticle.ID
		body_array("title")=objArticle.Title
		body_array("cateid")=objArticle.CateID
		body_array("tag")=objArticle.Tag
		body_array("content")=objArticle.Content
		body_array("level")=objArticle.Level
		body_array("authorid")=objArticle.AuthorID
		body_array("posttime")=objArticle.PostTime
		body_array("postcommnums")=objArticle.CommNums
		body_array("viewnums")=objArticle.ViewNums
		body_array("alias")=objArticle.Alias
		body_array("istop")=objArticle.Istop
		body_array("fullurl")=Replace(objArticle.FullUrl,"<#ZC_BLOG_HOST#>",BlogHost)
		body_array("type")=objArticle.FType
	Else
		errcode="111":msg="post id is wrong."
	End If
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取文章列表
' 参数: 
'*********************************************************
Function api_post_list
	Call VerifyApiKey()
	
	request_array(0)=Request.Form("PageSize")
	request_array(1)=Request.Form("Page")
	
	If (NOT IsNumeric(request_array(0))) OR (NOT IsNumeric(request_array(1))) Then errcode="111":msg="post id is wrong."
	
	If ErrorCheck() Then Exit Function

	Dim objRS,info_array,intPageAll,i
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""
	objRS.Open("SELECT * FROM [blog_Article] WHERE ([log_Type]=0) AND ([log_Level]>0) AND (1=1)  ORDER BY [log_PostTime] DESC")
	
	objRS.PageSize = request_array(0)
	If objRS.PageCount>0 Then objRS.AbsolutePage = request_array(1)
	intPageAll=objRS.PageCount
	
	If (Not objRS.bof) And (Not objRS.eof) Then
		For i=1 to objRS.PageSize
			Set info_array = jsObject()
			info_array("ID")=objRS("log_ID")
			info_array("CateID")=objRS("log_CateID")
			info_array("AuthorID")=objRS("log_AuthorID")
			info_array("Title")=objRS("log_Title")
			info_array("PostTime")=objRS("log_PostTime")
			info_array("CommNums")=objRS("log_CommNums")
			info_array("Level")=objRS("log_Level")
			Set body_array((i-1)) = info_array
			objRS.MoveNext
			If objRS.eof Then Exit For
		Next
	End If
	
	If  intPageAll>1 Then 
		data_export("pageall")=intPageAll
	End If 
	
	objRS.Close
	Set objRS=Nothing

	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取侧栏列表
' 参数: 
'*********************************************************
Function api_sidebar_list
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	Dim objRS,info_array,Doi:Doi=0
	Set objRS=objConn.Execute("SELECT [fn_ID],[fn_Name],[fn_FileName] FROM [blog_Function] ORDER BY [fn_ID] ASC")
	Do Until objRS.Eof
		Set info_array = jsObject()
		info_array("ID")=objRS("fn_ID")
		info_array("Name")=objRS("fn_Name")
		info_array("FileName")=objRS("fn_FileName")
		Set body_array(Doi) = info_array
		Doi=Doi+1
		objRS.MoveNext
	Loop
	
	Set objRS=Nothing
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取指定ID或者名称侧栏内容
' 参数: 
'*********************************************************
Function api_sidebar_get
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	request_array(0)=Request.Form("filename")
	
	Call GetFunction()
	
	Dim fnfilename,objFunction
	
	fnfilename=request_array(0)
	
	Set objFunction=New TFunction
	Set objFunction=Functions(FunctionMetas.GetValue(fnfilename))
	
	body_array("ID")=objFunction.ID
	body_array("Name")=objFunction.Name
	body_array("Order")=objFunction.Order
	body_array("Content")=objFunction.Content
	body_array("IsHidden")=objFunction.IsHidden
	body_array("SidebarID")=objFunction.SidebarID
	body_array("HtmlID")=objFunction.HtmlID
	body_array("Ftype")=objFunction.Ftype
	body_array("MaxLi")=objFunction.MaxLi
	body_array("Source")=objFunction.Source
	body_array("ViewType")=objFunction.ViewType

	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 添加tag
' 参数: 
'*********************************************************
Function api_tag_add
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	Dim objTag
	Set objTag=New TTag
	objTag.ID=Request.Form("ID")
	objTag.Name=Request.Form("Name")
	objTag.Intro=Request.Form("Intro")
	
	If CLng(objTag.ID)>0 Then 
		errcode="111":msg="tag add wrong."
	Else
		If objTag.Post Then
			Call GetTagsbyTagIDList("{"&objTag.ID&"}")
			body_array("add_type")=0
		End If
	End If
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 删除tag
' 参数: 
'*********************************************************
Function api_tag_del
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	Dim objTag
	Set objTag=New TTag
	objTag.ID=Request.Form("ID")

	Call GetTagsbyTagIDList("{"&objTag.ID&"}")

	If objTag.Del Then 
		body_array("del_type")=0
	Else
		errcode="111":msg="tag del wrong."
	End If
	Set objTag=Nothing

	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 修改tag
' 参数: 
'*********************************************************
Function api_tag_edit
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	Dim objTag
	Set objTag=New TTag
	objTag.ID=Request.Form("ID")
	objTag.Name=Request.Form("Name")
	objTag.Intro=Request.Form("Intro")

	If CLng(objTag.ID)>0 Then 
		objTag.MetaString=objConn.Execute("SELECT [tag_Meta] FROM [blog_Tag] WHERE [tag_ID]="&CLng(objTag.ID))(0)

		If objTag.Post Then
			Call GetTagsbyTagIDList("{"&objTag.ID&"}")
			body_array("edit_type")=0
		End If
	Else
		errcode="111":msg="tag edit wrong."
	End If
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取tag列表
' 参数: 
'*********************************************************
Function api_tag_list
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	Dim objRS,info_array,Doi:Doi=0
	Set objRS=objConn.Execute("SELECT [tag_ID],[tag_Name],[tag_Intro],[tag_Order],[tag_Count],[tag_ParentID],[tag_URL] FROM [blog_Tag] ORDER BY [tag_ID] ASC")
	Do Until objRS.Eof
		Set info_array = jsObject()
		info_array("ID")=objRS("tag_ID")
		info_array("Name")=objRS("tag_Name")
		info_array("Count")=objRS("tag_Count")
		info_array("Intro")=objRS("tag_Intro")
		Set body_array(Doi) = info_array
		Doi=Doi+1
		objRS.MoveNext
	Loop
	Set objRS=Nothing
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 修改用户信息
' 参数: 
'*********************************************************
Function api_user_edit
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function

	If (Request.Form("id"))="" Then
		errcode="888":msg="user id is empty"
	ElseIf Request.Form("id")<1 Then
		errcode="888":msg="user id is wrong"
	Else
		Dim objUser
		Set objUser=New TUser
		objUser.ID=Request.Form("ID")
		If (Request.Form("Level"))="" OR (Request.Form("Name"))="" OR (Request.Form("Email"))="" Then
			errcode="888":msg="Level、Name or Email is empty."
		End If
		objUser.Level=Request.Form("Level")
		objUser.Name=Request.Form("Name")
		objUser.Email=Request.Form("Email")
		objUser.HomePage=Request.Form("HomePage")
		objUser.Alias=Request.Form("Alias")
		objUser.Intro=Request.Form("Intro")
	End If
	
	If objUser.Edit(objUser) Then
		body_array("edit_type")=0
	Else
		errcode="888":msg="Level、Name or Email is empty."
	End IF
	Set objUser=Nothing
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取用户信息
' 参数: 
'*********************************************************
Function api_user_list
	Call VerifyApiKey()
	If ErrorCheck() Then Exit Function
	
	Dim objRS,info_array,Doi:Doi=0
	Set objRS=objConn.Execute("SELECT [mem_ID],[mem_Name],[mem_Level],[mem_Email],[mem_HomePage],[mem_PostLogs],[mem_Url],[mem_Intro],[mem_Meta] FROM [blog_Member]")
	Do Until objRS.Eof
		Set info_array = jsObject()
		info_array("ID")=objRS("mem_ID")
		info_array("Name")=objRS("mem_Name")
		info_array("Level")=objRS("mem_Level")
		info_array("Email")=objRS("mem_Email")
		info_array("HomePage")=objRS("mem_HomePage")
		info_array("PostLogs")=objRS("mem_PostLogs")
		info_array("Url")=objRS("mem_Url")
		info_array("Intro")=objRS("mem_Intro")
		info_array("Meta")=objRS("mem_Meta")
		Set body_array(Doi) = info_array
		Doi=Doi+1
		objRS.MoveNext
	Loop
	Set objRS=Nothing
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
End Function
%>