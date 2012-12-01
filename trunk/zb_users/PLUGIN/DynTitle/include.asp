<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    未寒
'//	http://imzhou.com
'//	挂口: 注册插件和接口
'///////////////////////////////////////////////////////////////////////////////




Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"" src=\""http://www.verycd.com/statics/title.saying\""></script>"");")

Call DynTitleTITLE()

'注册插件
Call RegisterPlugin("DynTitle","ActivePlugin_DynTitle")



' '接口
' Function ActivePlugin_DynTitle()

	
	' Call Add_Action_Plugin("Action_Plugin_TArticle_Export_Begin","Call DynTitleBlogTitle()")
	
	' Call Add_Action_Plugin("Action_Plugin_TArticleList_ExportByMixed_Begin","Call DynTitleBlogTitle()")
	
	' Call Add_Action_Plugin("Action_Plugin_Default_Begin","Call DynTitleTITLE()")

	' Call Add_Action_Plugin("Action_Plugin_TArticleList_Search_Begin","Call DynTitleTITLE()")

' End Function

' Function DynTitleBlogTitle()
	
	' Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"">document.title= \"""& BlogTitle & " – \""+ _VC_DocumentTitles[_VC_DocumentTitleIndex];</script>"");")

' End Function


Function DynTitleTITLE()
	
	Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","document.write(""<script type=\""text/javascript\"">document.title= \"""& ZC_BLOG_TITLE & " – \""+ _VC_DocumentTitles[_VC_DocumentTitleIndex];</script>"");")

End Function



%>