<%

'注册插件
Call RegisterPlugin("SuperCache","ActivePlugin_SuperCache")
'挂口部分
Function ActivePlugin_SuperCache()

	Call Add_Action_Plugin("Action_Plugin_Catalog_Begin","Dim SuperCache_Catch,SuperCache_Cache:SuperCache_Cache=SuperCache_Catalog():Call SuperCache_Export(SuperCache_Cache)")
	Call Add_Action_Plugin("Action_Plugin_Catalog_End","If SuperCache_Catch=False Then Call SuperCache_Save(SuperCache_Cache,ArtList.html)")
	
End Function

Function SuperCache_Save(FileName,Html)
	Call SaveToFile(BlogPath & "zb_users\cache\" & FileName & ".asp","<"&"%"&html,"utf-8",False)
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
	Dim s
	s=LoadFromFile(BlogPath & "zb_users\cache\" & FileName & ".asp","utf-8")
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

	Application.UnLock
	
	Dim fso,objFolder,objFile
	Set fso=Server.CreateObject("scripting.filesystemobject")
    Set objFolder=fso.GetFolder(BlogPath & "zb_users\cache")   
	
    For Each objFile in objFolder.Files
		If CheckRegExp(objFile.Name,"^supercache_.+?\.asp$") Then 
			Call DelToFile(objFile.Path)
		End If
    Next 
    Set objFolder=nothing   
    Set fso=nothing   

	ClearGlobeCache=True

End Function

%>