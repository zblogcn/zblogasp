<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    Cloudream
'//	http://labs.cloudream.name/z-blog/fronthelper/
'///////////////////////////////////////////////////////////////////////////////
Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"" src=\""" & BlogHost & "zb_users/plugin/FrontHelper/fronthelper.pack.js\""></script>"");")


'注册插件
Call RegisterPlugin("FrontHelper","ActivePlugin_FrontHelper")

Function ActivePlugin_FrontHelper()


	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Export_Template","FrontHelper_PostCodeInnerCode")

	Call Add_Filter_Plugin("Filter_Plugin_TComment_MakeTemplate_Template","FrontHelper_CommentCodeInnerCode")


End Function


'文章块模板代码插入
Function FrontHelper_PostCodeInnerCode(n,v)
	
	n = Replace(n,"<#article/content#>","<#article/content#><p style=""display:none;"" class=""cloudreamHelperLink"" codetype=""post"" entryid=""<#article/id#>""></p>")

	n = Replace(n,"<#article/intro#>","<#article/intro#><p style=""display:none;"" class=""cloudreamHelperLink"" codetype=""postmulti"" entryid=""<#article/id#>""></p>")

	v = Replace(v,"<#article/content#>","<#article/content#><p style=""display:none;"" class=""cloudreamHelperLink"" codetype=""post"" entryid=""<#article/id#>""></p>")

	v = Replace(v,"<#article/intro#>","<#article/intro#><p style=""display:none;"" class=""cloudreamHelperLink"" codetype=""postmulti"" entryid=""<#article/id#>""></p>")
	
	
End Function


'评论块模板代码插入
Function FrontHelper_CommentCodeInnerCode(ByRef html)

	html = Replace(html,"<#article/comment/content#>","<#article/comment/content#><p style=""display:none;"" class=""cloudreamHelperLink"" codetype=""comment"" entryid=""<#article/comment/id#>""></p>")
	
End Function



'安装插件
Function InstallPlugin_FrontHelper

	Call SetBlogHint(Empty,Empty,True)

End Function

'卸载插件
Function UnInstallPlugin_FrontHelper

	Call SetBlogHint(Empty,Empty,True)

End Function
%>