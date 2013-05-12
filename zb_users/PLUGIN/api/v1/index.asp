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

	Case "user_info"
		Call api_user_info()
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

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 删除分类
' 参数: 
'*********************************************************
Function api_cate_del

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 编辑分类
' 参数: 
'*********************************************************
Function api_cate_edit

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取分类列表
' 参数: 
'*********************************************************
Function api_cate_list

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 添加评论
' 参数: 
'*********************************************************
Function api_comment_add

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 删除评论
' 参数: 
'*********************************************************
Function api_comment_del

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 修改评论内容
' 参数: 
'*********************************************************
Function api_comment_edit

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取指定评论内容
' 参数: 
'*********************************************************
Function api_comment_get

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取评论列表
' 参数: 
'*********************************************************
Function api_comment_list

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
	
	request_array(0)=Request.Form("UserID")
	request_array(1)=Request.Form("autoname")
	
	Dim objUpLoadFile
	Set objUpLoadFile=New TUpLoadFile

	objUpLoadFile.AuthorID=request_array(0)
	objUpLoadFile.AutoName=request_array(1)
	
	If objUpLoadFile.UpLoad() Then
		UploadFile=True
		body_array("add_type")=0
	Else
		ret=1:errcode="002":msg="file upload wrong."
	End If
	Set objUpLoadFile=Nothing
	
	If ErrorCheck() Then Exit Function
	
	Set data_export("body")=body_array

	data_export.Flush
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
	
	Dim objRS,info_array,bodyID
	Set objRS=objConn.Execute("SELECT * FROM [blog_UpLoad] ORDER BY [ul_PostTime] DESC")
	Do Until objRS.Eof
		Set info_array = jsObject()
		info_array("upload_AuthorID")=objRS("ul_AuthorID")
		info_array("upload_FileSize")=objRS("ul_FileSize")
		info_array("upload_FileName")=objRS("ul_FileName")
		info_array("upload_PostTime")=objRS("ul_PostTime")
		info_array("upload_Meta")=objRS("ul_Meta")
		bodyID=objRS("ul_Id")
		Set body_array(bodyID) = info_array
		'Set info_array = Nothing
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

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 添加文章或独立页面
' 参数: 
'*********************************************************
Function api_post_add

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 删除指定ID文章或独立页面
' 参数: 
'*********************************************************
Function api_post_del

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 修改指定ID文章或独立页面
' 参数: 
'*********************************************************
Function api_post_edit

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取制定ID文章或独立页面
' 参数: 
'*********************************************************
Function api_post_get

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取文章列表
' 参数: 
'*********************************************************
Function api_post_list

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
	
	Dim objRS,info_array,bodyID
	Set objRS=objConn.Execute("SELECT [fn_ID],[fn_Name],[fn_FileName] FROM [blog_Function] ORDER BY [fn_ID] ASC")
	Do Until objRS.Eof
		Set info_array = jsObject()
		info_array("Name")=objRS("fn_Name")
		info_array("FileName")=objRS("fn_FileName")
		bodyID=objRS("fn_ID")
		Set body_array(bodyID) = info_array
		'Set info_array = Nothing
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
	
	Dim objRS,info_array,bodyID,Doi:Doi=0
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

End Function

'*********************************************************
' 作者: 未寒（im@imzhou.com）
' 修改: 2013-05-10
' 功能: 获取用户信息
' 参数: 
'*********************************************************
Function api_user_info

End Function

%>