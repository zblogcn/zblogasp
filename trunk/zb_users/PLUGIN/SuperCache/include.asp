<%

'注册插件
Call RegisterPlugin("SuperCache","ActivePlugin_SuperCache")
'挂口部分
Function ActivePlugin_SuperCache()

	Call Add_Action_Plugin("Action_Plugin_Catalog_Begin","Dim SuperCache_Catch,SuperCache_Cache:SuperCache_Cache=SuperCache_Catalog():Call SuperCache_Export(SuperCache_Cache)")
	Call Add_Action_Plugin("Action_Plugin_View_Begin","Dim SuperCache_Catch,SuperCache_Cache:SuperCache_Cache=SuperCache_View():Call SuperCache_Export(SuperCache_Cache)")
	Call Add_Action_Plugin("Action_Plugin_Tags_Begin","Dim SuperCache_Catch,SuperCache_Cache:SuperCache_Cache=SuperCache_Tags():Call SuperCache_Export(SuperCache_Cache)")


	Call Add_Action_Plugin("Action_Plugin_Catalog_End","If SuperCache_Catch=False Then Call SuperCache_Save(SuperCache_Cache,ArtList.html)")
	Call Add_Action_Plugin("Action_Plugin_View_End","If SuperCache_Catch=False Then Call SuperCache_Save(SuperCache_Cache,Article.html)")
	Call Add_Action_Plugin("Action_Plugin_Tags_End","If SuperCache_Catch=False Then Call SuperCache_Save(SuperCache_Cache,objArticle.html)")


End Function
Function SuperCache_Tags()
	SuperCache_Tags=Trim("supercache_tags")
End Function

Function SuperCache_View()
	SuperCache_View=Trim("supercache_view_"&vbsescape(Request.QueryString("id")))
End Function

Function SuperCache_Save(FileName,Html)
	Dim list
	list=Application(ZC_BLOG_CLSID&"SuperCache_Item")
	If Not IsArray(list) Then list=Array()
	Redim Preserve list(Ubound(list)+1)
	list(Ubound(list))=FileName
	Application.Lock()
	Application(ZC_BLOG_CLSID& "SuperCache_" & FileName)=Html
	Application(ZC_BLOG_CLSID&"SuperCache_Item")=list
	Application.Unlock()
End Function

Function SuperCache_Catalog()
	Dim aryFileName(4)
	aryFileName(0)=vbsescape(Request.QueryString("page"))
	aryFileName(1)=vbsescape(Request.QueryString("cate"))
	aryFileName(2)=vbsescape(auth)
	aryFileName(3)=vbsescape(Request.QueryString("date"))
	aryFileName(4)=vbsescape(Request.QueryString("tags"))
	SuperCache_Catalog=Trim("supercache_catalog_" & Join(aryFileName,"_"))
End Function


Function SuperCache_Export(FileName)
	Dim aryApt,i,s
	s=""
	aryApt=Application(ZC_BLOG_CLSID&"SuperCache_Item")
	If IsArray(aryApt) Then
		For i=0 To Ubound(aryApt)
			If aryApt(i)=FileName Then
				s=Application(ZC_BLOG_CLSID &"SuperCache_"&FileName)
				Exit For
			End If
		Next
	End If
	If Len(s)<>0 Then
		SuperCache_Catch=True
		Response.Write Replace(s,"<"&"%","",1,1)
		Response.Write vbCrlf & "<!--Export From Cache-->"
		Response.Write vbCrlf & "<!-- " & RunTime() & "ms -->"
		Response.End
	Else
		SuperCache_Catch=False
	End If
End Function



'重写系统函数

Function ClearGlobeCache()

	Application.Lock


	Application(ZC_BLOG_CLSID & "TemplateTagsName")=Empty
	Application(ZC_BLOG_CLSID & "TemplateTagsValue")=Empty

	Application(ZC_BLOG_CLSID & "TemplatesName")=Empty
	Application(ZC_BLOG_CLSID & "TemplatesContent")=Empty


	Application(ZC_BLOG_CLSID & "SIGNAL_RELOADCACHE")=Empty

	Application(ZC_BLOG_CLSID & "TEMPLATEMODIFIED")=Empty

	Application(ZC_BLOG_CLSID & "CACHE_ARTICLE_VIEWCOUNT")=Empty
	
	Application(ZC_BLOG_CLSID & "SuperCache_Item")=Empty

	Application.UnLock
	
	

	ClearGlobeCache=True

End Function

%>