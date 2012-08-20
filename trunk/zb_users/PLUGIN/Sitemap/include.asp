<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    sipo / Cloudream
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    sitemap.asp
'// 开始时间:    2006-8-3
'// 最后修改:    
'// 备    注:    
'//	http://labs.cloudream.name/z-blog/Sitemap/
'///////////////////////////////////////////////////////////////////////////////

'静态文件名
Const sitemapFileName = ""
Const sitemapWAPFileName = "sitemap_wap.xml"

'注册插件
Call RegisterPlugin("Sitemap","ActivePlugin_Sitemap")
Dim Sitemapc
Function SiteMap_Initialize
	Set Sitemapc=New TConfig
	Sitemapc.Load "SiteMap"
	If Sitemapc.Exists("v")=False Then
		Sitemapc.Write "v","1.0"
		Sitemapc.Write "c","0"
		Sitemapc.Save
	End If
End Function

Function ActivePlugin_Sitemap()
	Call Add_Action_Plugin("Action_Plugin_ArticlePst_Succeed","Call ExportSiteMap")
	Call Add_Action_Plugin("Action_Plugin_MakeFileReBuild_End","Call ExportSiteMap")
End Function

Function SiteMap
	Call AddBatch("生成SiteMap","ExportSiteMap")	
End Function

'安装插件
Function InstallPlugin_Sitemap
	SiteMap
	SiteMap_Initialize
End Function


Function ExportSiteMap
	Dim a
	Set a=New SiteMap_Export
	a.Export
	Set a=Nothing
End Function

Class SiteMap_Export
'去你妹的XML！
	Public TimeZone
	Dim FType
	Dim xmlDom,xmlwapDom
	
	Public Property Get xml
		xml = xmlDom.xml
	End Property
	
	Public Property Get wapxml
		wapxml = xmlwapDom.xml
	End Property
	
	Dim cxml,wapcxml
	'**************************************************'
	'                     类初始化                       '
	'**************************************************'
	Sub Class_Initialize()
		TimeZone=ZC_TIME_ZONE		
		Set xmlDom =Server.CreateObject("Microsoft.XMLDOM")
		xmlDom.insertBefore xmlDom.createProcessingInstruction("xml","version=""1.0"" encoding=""utf-8"""), xmlDom.childNodes(0)
		Set xmlwapDom=Server.CreateObject("Microsoft.XMLDOM")
		xmlwapDom.insertBefore xmlwapDom.createProcessingInstruction("xml","version=""1.0"" encoding=""utf-8"""), xmlwapDom.childNodes(0)
		'初始化XMLDOM
		Set cxml=xmlDom.createElement("urlset")
		xmlDom.AppendChild(cxml)
		cxml.setAttribute "xmlns","http://www.Sitemap.org/schemas/sitemap/0.9"	
		'初始化WEB SITEMAP
		Set wapcxml=xmlWapDom.createElement("urlset")
		xmlwapDom.AppendChild(wapcxml)
		wapcxml.setAttribute "xmlns:mobile","http://www.google.com/schemas/sitemap-mobile/1.0"
		wapcxml.setAttribute "xmlns","http://www.Sitemap.org/schemas/sitemap/0.9"	
		'初始化 WAP SITEMAP
	End Sub
	'**************************************************'
	'                 空白转换成EMPTY                    '
	'**************************************************'
	Sub toEmpty(ByRef str)
		If str="" Or IsNull(str) Or IsEmpty(str) Then str=Empty
	End Sub
	'**************************************************'
	'                  新建XML节点                       '
	'**************************************************'
	Public Function Add(url,timestamp,changefreq,priority,wapurl)
		'VBS居然不支持Optional！纠结！
		'看看JS……
		Dim o,wo
		toEmpty url
		toEmpty timestamp
		toEmpty changefreq
		toEmpty priority
		toEmpty wapurl
		
		Set o = xmlDom.createElement("url")
		Set wo= xmlWapDom.createElement("url")
		If Not IsEmpty(url) Then
			o.AppendChild(xmlDom.createElement("loc"))
			o.selectSingleNode("loc").text=url
		End If
		If Not IsEmpty(wapurl) Then
			wo.AppendChild(xmlWapDom.createElement("loc"))
			wo.selectSingleNode("loc").text=wapurl
			wo.AppendChild(xmlWapDom.createElement("mobile:mobile"))
			wapcxml.AppendChild(wo)
		End If
		If Not IsEmpty(timestamp) Then
			o.AppendChild(xmlDom.createElement("lastmod"))
			o.selectSingleNode("lastmod").text=GetDate(timestamp)
		End If
		If Not IsEmpty(changefreq) Then
			o.AppendChild(xmlDom.createElement("changefreq"))
			o.selectSingleNode("changefreq").text=changefreq
		End If
		If Not IsEmpty(priority) Then
			o.AppendChild(xmlDom.createElement("priority"))
			o.selectSingleNode("priority").text=priority
		End If
		cxml.AppendChild(o)

	End Function
	'**************************************************'
	'                得到那个奇怪的日期                   '
	'**************************************************'
	Function GetDate(dat)
		Dim d,w,m,y
		Dim h,mi,s,z
		y = ToTwo(dat,"y")
		m = ToTwo(dat,"m")
		d = ToTwo(dat,"d")
		h = ToTwo(dat,"h")
		s = ToTwo(dat,"s")
		mi= ToTwo(dat,"mi")
		z=Left(TimeZone,3) & ":" & Right(TimeZone,2)
		GetDate = y & "-" & m & "-" & d & "T" & h & ":" & mi & ":" & s & z
	End Function 
	'**************************************************'
	'                 一位数转二位数                     '
	'**************************************************'
	Function ToTwo(i,s)
		dim j
		j=i
		Select Case s
			Case "m":j=Month(j)
			Case "d":j=Day(j)
			Case "mi":j=Minute(j)
			Case "s":j=Second(j)
			Case "h":j=Hour(j)
			Case "y":j=Year(j)
		End Select
		If Len(j)=1 Then j="0"&j
		ToTwo=j
	End Function
	'**************************************************'
	'                        输出                       '
	'**************************************************'
	Function Export
	
		Add ZC_BLOG_HOST,Empty,"daily",Empty,ZC_BLOG_HOST&"wap.asp"
		'输出博客地址
		Dim i,URL,objRS
	
		Dim objArticle
		Dim C
		Call SiteMap_Initialize
		C=SiteMapC.Read("c")
		If c="" Then c=0
		c=CInt(c)
		'读取配置
		Set objRS=objConn.Execute("SELECT [log_ID],[log_Tag],[log_CateID],[log_Level],[log_AuthorID],[log_PostTime],[log_Istop],[log_Url],[log_Type] FROM [blog_Article] WHERE ([log_ID]>0) AND ([log_Level]>2) AND ([log_Type]="&ZC_POST_TYPE_ARTICLE&") ORDER BY [log_PostTime] DESC")
		'读取数据库

		Do Until objRs.Eof	
			
			If i=C And c>0 then Exit Do
			Set objArticle=New TArticle

			If objArticle.LoadInfoByArray(Array(objRs(0),objRs(1),objRs(2),"","","",objRs(3),objRs(4),objRs(5),0,0,0,objRs(7),objRs(6),"","",objRs(8),"")) Then
				Add objArticle.Url,objArticle.PostTime,"monthly",IIf(objArticle.IsTop=True,"1.0",Empty),objArticle.WapUrl
				'添加数据
			End If
			i=i+1
			Set objArticle=Nothing
			objRS.MoveNext
		Loop
		objRS.Close
		Set objRS=Nothing
		Call SaveToFile(BlogPath&"/sitemap.xml",xml,"UTF-8",False)
		Call SaveToFile(BlogPath&"/sitemap_wap.xml",wapxml,"UTF-8",False)
		'保存
	End Function

End Class

%>


