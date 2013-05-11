<%
Function Public_date
	data_export("errcode") = errcode
	data_export("msg") = msg
	data_export("ret") = ret
	data_export("timestamp") = DateDiff("s", "1970/01/01 00:00:00", Now())
End Function

Function VerifyApiKey
	If Request("keyid")<>objConfig.Read("id") Then
		errcode="001":msg="keyid is wrong."
	ElseIf Request("keysecretmd5")<>md5(objConfig.Read("secret")) Then
		errcode="002":msg="keysecret is wrong."
	End If
End Function	

Function Verify()
	Call VerifyApiKey()
	If errcode<>"0" Then Exit Function
	
	request_array(0) = Request("post")
	
	If request_array(0)<>"" Then
		body_array("blog_title")=ZC_BLOG_TITLE
		body_array("blog_subtitle")=ZC_BLOG_SUBTITLE
		body_array("blog_url")=BlogHost
		body_array("blog_language")=ZC_BLOG_LANGUAGE
		body_array("blog_version")=BlogVersion
	Else
		errcode="003":msg="post data is empty."
	End If
End Function

Function GetBlogSet
	body_array("blog_mssqlenable")=ZC_MSSQL_ENABLE
	body_array("blog_using_plugin")=ZC_USING_PLUGIN_LIST
	body_array("blog_userzone")=ZC_TIME_ZONE
	body_array("blog_hostzone")=ZC_HOST_TIME_ZONE
	body_array("blog_multidomain")=ZC_MULTI_DOMAIN_SUPPORT
	body_array("blog_version_name")=ZC_BLOG_VERSION
	body_array("blog_commentoff")=ZC_COMMENT_TURNOFF
	body_array("blog_commnetfloor")=ZC_COMMNET_MAXFLOOR
	body_array("blog_display_count")=ZC_DISPLAY_COUNT
	body_array("blog_rss_count")=ZC_RSS2_COUNT
	body_array("blog_search_count")=ZC_SEARCH_COUNT
	body_array("blog_pagebar_count")=ZC_PAGEBAR_COUNT
	body_array("blog_mutuality_count")=ZC_MUTUALITY_COUNT
	body_array("blog_comments_count")=ZC_COMMENTS_DISPLAY_COUNT
	body_array("blog_rss_whole")=ZC_RSS_EXPORT_WHOLE
	body_array("blog_template_die")=ZC_TEMPLATE_DIRECTORY
	body_array("blog_upload_type")=ZC_UPLOAD_FILETYPE
	body_array("blog_upload_size")=ZC_UPLOAD_FILESIZE
	body_array("blog_upload_dir")=ZC_UPLOAD_DIRECTORY
	body_array("blog_static_type")=ZC_STATIC_TYPE
	body_array("blog_static_dir")=ZC_STATIC_DIRECTORY
	body_array("blog_static_mode")=ZC_POST_STATIC_MODE
	body_array("blog_article_link")=ZC_ARTICLE_REGEX
	body_array("blog_page_link")=ZC_PAGE_REGEX
	body_array("blog_cate_link")=ZC_CATEGORY_REGEX
	body_array("blog_page_link")=ZC_PAGE_REGEX
End Function
%>