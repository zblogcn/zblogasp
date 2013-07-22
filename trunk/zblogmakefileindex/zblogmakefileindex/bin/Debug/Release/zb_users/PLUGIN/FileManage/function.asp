<%
Dim FileManage_FSO


'*********************************************************
' 目的：    格式化文件大小
'*********************************************************
Function FileManage_GetSize(FileSize)
	For Each sAction_Plugin_FileManage_GetSize_Begin In Action_Plugin_FileManage_GetSize_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_GetSize_Begin) Then Call Execute(sAction_Plugin_FileManage_GetSize_Begin)
	Next
	
	Dim b,m
	b=filesize:m="B"
	if b>1024 then b=b/1024:m="K"
	if b>1024 then b=b/1024:m="M"
	if b>1024 then b=b/1024:m="G"
	b=formatnumber(b,2)
	FileManage_GetSize=b&m

	For Each sAction_Plugin_FileManage_GetSize_End In Action_Plugin_FileManage_GetSize_End
		If Not IsEmpty(sAction_Plugin_FileManage_GetSize_End) Then Call Execute(sAction_Plugin_FileManage_GetSize_End)
	Next
End Function

'*********************************************************
' 目的：    得到文件图标
'*********************************************************
Function FileManage_GetTypeIco(FileName)
	For Each sAction_Plugin_FileManage_GetTypeIco_Begin In Action_Plugin_FileManage_GetTypeIco_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_GetTypeIco_Begin) Then Call Execute(sAction_Plugin_FileManage_GetTypeIco_Begin)
	Next
	
	Dim aryFn
	aryFn=Split(FileName,".")
	Dim strType
	strType=LCase(aryFn(UBound(aryFn)))
	Dim ImgTag,Tag
	ImgTag="<img width=""16"" height=""16"" src=""..\..\..\zb_system\IMAGE\FILETYPE\{tag}.png""/>"
	Select Case strType
		Case "jar","jad" Tag="jar"
		Case "txt","config","ini","inf","log" Tag="txt"
		Case "doc","docx","docm","dot","dotx","docm","odt","wpd","rtf","wps" Tag="doc"
		Case "ppt","pptx","pptm","ppsx","pps","ppsm","potx","pot","potm","odp" Tag="ppt"
		Case "xls","xlsm","xlsb","xl","xlam","xltc","xltm","xla","odc","ods" Tag="xls"
		Case "pdf" Tag="pdf"
		Case "sql" Tag="sql"
		Case "mp3","wma","wav","ogg" Tag="msc"
		Case "mpg","mpeg","avi","rm","rmvb","vob","dat","mp4","3gp","flv","swf","mkv","mov" Tag="mov"
		Case "exe","com" Tag="exe"
		Case "dll","ocx","sys","db" Tag="dll"
		Case "bat","cmd" Tag="bat"
		Case "asp","php","jsp","js","css","inc","asa","asax","aspx","mhtml","shtml","py"  Tag="code"
		Case "jpg","jpeg","gif","bmp","png","tiff","ico" Tag="img"
		Case "htm","html","xml"  Tag="htm"
		Case "rar","zip","7z","gz"  Tag="rar"
		Case "mdb" Tag="acc"
		Case "zba","zti","zpi" Tag="zba"

		Case Else  		
			Dim strFound
			For Each sAction_Plugin_FileManage_GetTypeIco_NotFound In Action_Plugin_FileManage_GetTypeIco_NotFound
				If Not IsEmpty(sAction_Plugin_FileManage_GetTypeIco_NotFound) Then
					sAction_Plugin_FileManage_GetTypeIco_NotFound=Replace(Replace(sAction_Plugin_FileManage_GetTypeIco_NotFound,"{path}",Replace(path,"""","""""")),"{f}",Replace(foldername,"""",""""""))
					Execute "strFound="&sAction_Plugin_FileManage_GetTypeIco_NotFound&vbcrlf&"if strFound<>"""" then Tag=strfound"
				End If
			Next
			If Tag="" Then Tag="no"
	End Select
	FileManage_GetTypeIco=Replace(ImgTag,"{tag}",tag)

	For Each sAction_Plugin_FileManage_GetTypeIco_End In Action_Plugin_FileManage_GetTypeIco_End
		If Not IsEmpty(sAction_Plugin_FileManage_GetTypeIco_End) Then Call Execute(sAction_Plugin_FileManage_GetTypeIco_End)
	Next
End Function
'*********************************************************
' 目的：    输出注释
'*********************************************************

Function FileManage_ExportInformation(foldername,path)
	For Each sAction_Plugin_FileManage_ExportInformation_Begin In Action_Plugin_FileManage_ExportInformation_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_ExportInformation_Begin) Then Call Execute(sAction_Plugin_FileManage_ExportInformation_Begin)
	Next

	Dim z,k,l,n
	n=""
	
	z=LCase(foldername)
	k=LCase(path)
	If foldername="" Then
		n=Split(k,"\")
		z=LCase(n(UBound(n)))
		ReDim Preserve n(UBound(n)-1)
		k=LCase(Join(n,"\"))
		n=""
	End If
	l=LCase(BlogPath)
	
	l=IIf(Right(l,1)="\",Left(l,Len(l)-1),l)
	k=IIf(Right(k,1)="\",Left(k,Len(k)-1),k)
	

	
	

	Dim h
	h=Replace(LCase(ZC_DATABASE_PATH),"/","\")

	if k=l then
	
		Select Case z
			Case "zb_system" n="Z-Blog系统核心文件"
			Case "zb_users" n="Z-Blog用户配置文件夹"
			Case "zb_install" n="Z-Blog安装文件夹"
			Case LCase(ZC_STATIC_DIRECTORY) n="静态文件存放文件夹"
			Case "catalog.asp" n="文章列表"
			Case "default.asp" n="首页"
			Case "feed.asp","rss.xml" n="RSS订阅"
			Case "search.asp" n="搜索"
			Case "tags.asp" n="Tags列表"
			Case "wap.asp" n="Wap"
			Case "robots.txt" n="Robots"
			Case "sitemap.xml" n="站点地图"
			Case "view.asp" n="动态文章浏览"

		End Select
	ElseIf k&"\"&z=l&"\"&LCase(ZC_UPLOAD_DIRECTORY) then
		n="上传文件夹"
	ElseIf k&"\"&z=l&"\"&h then
		n="当前数据库"
	ElseIf k=l & "\zb_system" then
		Select Case z
			Case "admin" n="Z-Blog管理文件"
			Case "css" n="Z-Blog后台CSS文件夹"
			Case "function" n="核心文件"
			Case "image" n="后台资源文件夹"
			Case "script" n="后台脚本文件夹"
			Case "wap" n="Wap组件"
			Case "xml-rpc" n="Xml-Rpc组件"
			Case "defend" n="默认调用文件夹"
			Case "cmd.asp" n="命令执行跳转"
			Case "login.asp" n="登录"
		End Select
	ElseIf k=l & "\zb_users" then
		Select Case z
			Case "avatar" n="头像缓存文件夹"
			Case "cache" n="缓存文件夹"
			Case "data" n="数据库存放位置"
			Case "include" n="Z-Blog引用文件夹"
			Case "language" n="Language Pack"
			Case "plugin" n="插件文件夹"
			Case "theme" n="主题文件夹"
			Case "emotion" n="表情文件夹"
			'Case Replace(LCase(ZC_UPLOAD_DIRECTORY),"zb_users\") n="上传文件文件夹"
			Case "c_option.asp" n="网站设置文件"
		End Select
	ElseIf k=l &  "\zb_users\include" then
		Select Case z
			Case "link.asp" n="友链"
			Case "favorite.asp" n="收藏"
			Case "navbar.asp" n="导航栏"
			Case "misc.asp" n="图标汇总"			
		End Select
	ElseIf k=l &  "\zb_users\data" then
			'if CheckRegExp(z,".+?mdb|.+?asp") then n="可能是Z-Blog数据库"
	ElseIf k=l & "\zb_users\theme\" & LCase(ZC_BLOG_THEME) then 
		Select Case z
			Case "include" n="引用"
			Case "plugin" n="主题自带插件"
			Case "source" n="主题CSS"
			Case "style" n="主题CSS"
			Case LCase(ZC_TEMPLATE_DIRECTORY) n="主题模板"
		End Select
	ElseIf k=l & "\zb_users\theme\"&LCase(zc_blog_theme)&"\"&LCase(ZC_TEMPLATE_DIRECTORY) then
		z=Split(z,".")(0)
		Select Case z
			Case "b_article-istop" n= "首页置顶文章模板"
			Case "b_article-multi" n= "首页摘要文章模板"
			Case "b_article-single" n= "日志页文章模板"
			'Case "b_article-guestbook" n= "留言页正文模板"
			Case "b_article_comment" n= "每条评论内容显示模板"
			Case "b_article_commentrev" n="回复的评论显示模板"
			Case "b_article_commentpost-verify" n= "评论验证码显示样式"
			Case "b_article_commentpost" n= "评论发表框模板"
			Case "b_article_mutuality" n= "每条相关文章显示模板"
			Case "b_article_nvabar_l" n= "“上一篇”日志链接"
			Case "b_article_nvabar_r" n= "“下一篇”日志链接"
			Case "b_article_tag" n="Tag显示样式"
			Case "b_article-page" n="独立页面内容模板"
			Case "b_pagebar" n="分页条模板"
			Case "b_function" n="单个侧边栏模板"
			Case "b_article_comment_pagebar" n="评论分页模板"
			Case "catalog" n="分类页整页模板"
			Case "default" n="首页整页模板"
			Case "page" n="独立页面模板"
			Case "single" n="日志页整页模板"
			Case "header" n="头部模板"
			Case "footer" n="底部模板"
		End Select
	'ElseIf k=l & "\zb_users\theme\"&LCase(zc_blog_theme)&"\include" then
		'n="<#TEMPLATE_INCLUDE_"&uCase(Split(z,".")(0))&"#>"
	ElseIf k=l &"\zb_system\admin" then
		Select Case z
			Case "admin.asp" n="管理页"
			Case "admin_default.asp" n="主面板"
			Case "admin_left.asp" n="左侧面板"
			Case "admin_top.asp" n="后台头文件"
			Case "c_autosaverjs.asp" n="自动保存"
			Case "c_updateinfo.asp" n="得到最新消息"
			Case "c_statistic.asp" n="后台统计数据"
			Case "edit_catalog.asp" n="编辑分类页"
			Case "edit_comment.asp" n="编辑评论页"
			Case "edit_link.asp" n="链接管理页"
			Case "edit_setting.asp" n="网站设置页"
			Case "edit_tag.asp" n="Tag修改页"
			Case "edit_ueditor.asp" n="新建文章页"
			Case "edit_user.asp" n="用户编辑页"
			Case "ueditor" n="uEditor主文件"
			Case "admin_footer.asp" n="后台底部引用文件"
			Case "admin_header.asp" n="后台头部引用文件"
			Case "edit_function.asp" n="编辑侧栏页"
		End Select
	ElseIf k=l & "\zb_system\admin\ueditor" then
		Select Case z
			Case "asp" n="uEditor ASP后台"
			Case "dialogs" n="uEditor 对话框"
			Case "themes" n="uEditor 主题"
			Case "lang" n="uEditor 语言包"
			Case "third-party" n="第三方组件"
			Case "editor_all_min.js" n="uEditor主文件（压缩）"
			Case "editor_all.js" n="uEditor主文件"
			Case "editor_config.asp" n="uEditor配置"
		End Select
	ElseIf k=l & "\zb_system\admin\ueditor\asp" then
		Select Case z
			Case "fileup.asp" n="文件上传"
			Case "getmovie.asp" n="视频搜索"
			Case "getremoteimage.asp" n="下载远程图片"
			Case "imagemanager.asp" n="图片管理"
			Case "imageup.asp" n="图片上传"
			Case "aspincludefile.asp" n="上传需要函数"
			Case "scrawlup.asp" n="涂鸦上传"
		End Select
	ElseIf k=l & "\zb_system\function" then
		Select Case z
			Case "c_error.asp" n="Z-Blog错误处理"
			Case "c_function.asp" n="Z-Blog一般函数"
			Case "c_html_js.asp" n="访问计数等JS调用"
			Case "c_html_js_add.asp" n="动态JS调用文件"
			Case "c_system_base.asp" n="Z-Blog基础"
			Case "c_system_event.asp" n="Z-Blog事件"
			Case "c_system_lib.asp" n="Z-Blog 数据库访问类"
			Case "c_system_manage.asp" n="Z-Blog 后台管理文件"
			Case "c_system_plugin.asp" n="Z-Blog 插件支持文件"
			Case "c_system_wap.asp" n="Z-Blog Wap支持文件"
			Case "c_urlredirect.asp" n="Z-Blog 加密Url跳转页"
			Case "c_validcode.asp" n="Z-Blog验证码"
			Case "c_admin_js_add.asp","c_admin_js.asp" n="后台JS调用文件"
		End Select
	ElseIf k=l & "\zb_system\wap" then
		Select Case z
			Case "default.asp" n="Wap首页"
			Case "index.asp" n="Wap首页"
			Case "style" n="WapCSS"
			Case "wap_article-multi.html" n="Wap模板-文章"
			Case "wap_article_comment.html" n="Wap模板-评论"
			Case "wap_single.html" n="Wap模板-文章页或列表页"
		End Select
	
	else
		Dim strFound
		For Each sAction_Plugin_FileManage_ExportInformation_NotFound In Action_Plugin_FileManage_ExportInformation_NotFound
			If Not IsEmpty(sAction_Plugin_FileManage_ExportInformation_NotFound) Then
				sAction_Plugin_FileManage_ExportInformation_NotFound=Replace(Replace(sAction_Plugin_FileManage_ExportInformation_NotFound,"{path}",Replace(path,"""","""""")),"{f}",Replace(foldername,"""",""""""))
				Execute "strFound="&sAction_Plugin_FileManage_ExportInformation_NotFound&vbcrlf&"if strFound<>"""" then n=strfound"
			End If
		Next
	end if
	For Each sAction_Plugin_FileManage_ExportInformation_End In Action_Plugin_FileManage_ExportInformation_End
		If Not IsEmpty(sAction_Plugin_FileManage_ExportInformation_End) Then Call Execute(sAction_Plugin_FileManage_ExportInformation_End)
	Next
	FileManage_ExportInformation=n
End Function
'*********************************************************
' 目的：    输出文件列表
'*********************************************************
Function FileManage_ExportSiteFileList(path,OpenFolderPath)
	For Each sAction_Plugin_FileManage_ExportSiteFileList_Begin In Action_Plugin_FileManage_ExportSiteFileList_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteFileList_Begin) Then Call Execute(sAction_Plugin_FileManage_ExportSiteFileList_Begin)
	Next
	

	'On Error Resume Next
	Dim f,fold,item,fpath,jpath
	If OpenFolderPath<>"" Then path=OpenFolderPath

	FileManage_FormatPath path
	Dim backfolder
	backfolder=Split(path,"\")
	ReDim Preserve backfolder(UBound(backfolder)-1)
	backfolder=Join(backfolder,"\")
	If FileManage_CheckFolder(path) Then Response.Write  "<p>当前路径:" & path & "</p><p>对不起，为了您的其他程序的安全，您只能修改Z-Blog文件夹内的文件，同时也不允许修改Global.asa和Global.asax。</p><p><a href='main.asp?act=SiteFileMng&path="&Server.URLEncode(BlogPath)&"'>点击这里返回</a></p></div>" :Response.end
	Set f=Server.Createobject("scripting.filesystemobject")


	Response.Write "<div id=""fileUpload"">"
	FileManage_ExportSiteUpload(path)
	Response.Write "</div>"
	Response.Write "<p>"&ZC_MSG240&":"&path&"</p>"
		If Instr(LCase(path),LCase(blogpath)&"zb_system")>0  then 
			Response.Write "<p><font color=""red"">注意！您正在使用的Z-Blog版本为"&ZC_BLOG_VERSION&"，修改系统文件请小心！</font></p>"
		End If
	
	
	Set fold=f.getfolder(path)
	
	Response.Write "<table width=""100%"" border=""0"" class=""tableBorder"">"
	Response.Write "<tbody>"
	Response.Write "<tr><th colspan=""5""><a h='_' href='main.asp?act=SiteFileMng&path="&Server.URLEncode(backfolder)&"' title='"&ZC_MSG239&"'><img src=""images\up.png"" width=""16"" alt="""" /></a>"
	Response.Write "&nbsp;&nbsp;<a href=""javascript:void(0)"" onclick=""$('#fileUpload').dialog({width:'741px',title:'上传'})"" title=""上传""><img src=""images\upload.png"" width=""16"" alt="""" /></a>"
'	Response.Write "&nbsp;&nbsp;<a href='javascript:void(0)' onclick='window.open(""main.asp?act=SiteFileUploadShow&path="&Server.URLEncode(fpath)&"&OpenFolderPath="& Server.URLEncode(path) &""",""Detail"",""Scrollbars=no,Toolbar=no,Location=no,Direction=no,Resizeable=no,height=165px,width=780px"")' title=""上传""><img src=""images\upload.png""/></a>"
	Response.Write "&nbsp;&nbsp;<a href='main.asp?act=SiteCreateFolder' h='_' onmousedown=""var str=prompt('请输入文件夹名');if(str!=null){this.href+='&path='+encodeURIComponent('"&Replace(Replace(path,"\","\\"),"""","\""")&"'+'\\'+str);this.click()}else{return false}"" title='新建文件夹'><img src='images\cfolder.png' width='16' alt='' /></a>"
	Response.Write "&nbsp;&nbsp;<a h='_' href=""main.asp?act=SiteFileEdt&path="&Server.URLEncode(path) &"&OpenFolderPath="&Server.URLEncode(path)&""" title=""创建文件""><img src=""images\newfile.png"" width=""16"" alt="""" /></a>"

	
	
	For Each sAction_Plugin_FileManage_AddControlBar In Action_Plugin_FileManage_AddControlBar
		If Not IsEmpty(sAction_Plugin_FileManage_AddControlBar) Then Call Execute(sAction_Plugin_FileManage_AddControlBar)
	Next
	
	
	
	Response.Write "</th></tr></tbody></table>"
	Response.Write "<table width=""100%"" border=""0"" class=""tableBorder"" id=""siteList"">"

	Response.Write "<thead><tr><th>文件名</th><th width=""17%"">修改时间</th><th width=""7%"">大小</th><th width=""24%"">注释</th><th>操作</th></tr></thead><tbody>"
	
	
	For Each item In fold.subfolders
	
		fpath=path&"/"&item.name
		fpath=Replace(Replace(fpath,"/","\"),"\\","\")
		jpath=Replace(path,"\","\\")
		Response.Write "<tr height='14'><td><img width=""16"" height=""16"" src='../../../zb_system/IMAGE/FILETYPE/folder.png' />&nbsp;<a href='main.asp?act=SiteFileMng&path="&Server.URLEncode(path&IIf(Right(path,1)="\","","\")&item.name)&"&OpenFolderPath=' h='_'>"&item.name&"</a>"
		Response.Write"</td><td>"&FormatDateTime(item.datelastmodified,0)&"</td><td></td><td>"&FileManage_ExportInformation(item.name,path)&"</td><td width=""15%"" align=""center"">"
		Response.Write "&nbsp;&nbsp;<a h='_' href=""main.asp?act=SiteFileRename&path="&Server.URLEncode(fpath)&"&OpenFolderPath="& Server.URLEncode(path) &""" class='rename_folder' title=""[重命名]""><img src="""&ZC_BLOG_HOST&"/zb_system/image/admin/document-rename.png"" width=""16"" height=""16"" alt='重命名' title='重命名'/></a>&nbsp;"
		Response.Write "&nbsp;&nbsp;<a h='_' href=""main.asp?act=SiteFileDel&folder=true&path="&Server.URLEncode(fpath)&"&OpenFolderPath="& Server.URLEncode(path) &""" class='delete_folder' title=""["&ZC_MSG063&"]""><img src="""&ZC_BLOG_HOST&"/zb_system/image/admin/delete.png"" width=""16"" height=""16"" alt='删除' title='删除'/></a>"
		Response.Write "</td></tr>"
	Next
	
	
	For Each item In fold.files
		fpath=path&"/"&item.name
		fpath=Replace(Replace(fpath,"/","\"),"\\","\")
		Response.Write "<tr><td>"&FileManage_GetTypeIco(item.name)&"&nbsp;<a href="""
		Dim isEmptyPlugin
		isEmptyPlugin=True
		For Each sAction_Plugin_FileManage_FileOpenType In Action_Plugin_FileManage_FileOpenType
			If Not IsEmpty(sAction_Plugin_FileManage_FileOpenType) Then
				Call Execute(sAction_Plugin_FileManage_FileOpenType)
				isEmptyPlugin=False
			End If
		Next
		If isEmptyPlugin Then Response.Write BlogHost & Replace(Replace(LCase(path)&"\",LCase(blogpath),"")&"/"&item.name,"//","/")
		
		Response.Write """ target=""_blank"" title='"&FormatDateTime(item.datelastmodified,0)&";"&ZC_MSG238&":"&clng(item.size/1024)&"k'>"&item.name&"</a></td><td>"&FormatDateTime(item.datelastmodified,0)&"</td><td>"&FileManage_GetSize(item.size)&"</td><td>"&FileManage_ExportInformation(item.name,path)&"</td><td align=""center"">"
		Response.Write"<a h='_' href=""main.asp?act=SiteFileEdt&path="&Server.URLEncode(fpath)&"&OpenFolderPath="& Server.URLEncode(path) &""" title=""["&ZC_MSG078&"]""><img src="""&ZC_BLOG_HOST&"/zb_system/image/admin/script_edit.png"" width=""16"" height=""16"" alt='编辑' title='编辑'/></a>&nbsp;"
		Response.Write "&nbsp;&nbsp;<a href=""main.asp?act=SiteFileDownload&path="&Server.URLEncode(fpath)&"&OpenFolderPath="& Server.URLEncode(path) &""" target=""_blank"" title=""[下载]""><img src="""&ZC_BLOG_HOST&"/zb_system/image/admin/download.png"" width=""16"" height=""16"" alt='下载' title='下载'/></a>&nbsp;"
		Response.Write "&nbsp;&nbsp;<a h='_' href=""main.asp?act=SiteFileRename&path="&Server.URLEncode(fpath)&"&OpenFolderPath="& Server.URLEncode(path) &""" class=""rename_file"" title=""[重命名]""><img src="""&ZC_BLOG_HOST&"/zb_system/image/admin/document-rename.png"" width=""16"" height=""16"" alt='重命名' title='重命名'/></a>&nbsp;"
	
		Response.Write "&nbsp;&nbsp;<a h='_' href=""main.asp?act=SiteFileDel&path="&Server.URLEncode(fpath)&"&OpenFolderPath="& Server.URLEncode(path) &""" class=""delete_file"" title=""["&ZC_MSG063&"]""><img src="""&ZC_BLOG_HOST&"/zb_system/image/admin/delete.png"" width=""16"" height=""16"" alt='删除' title='删除'/></a>"
		For Each sAction_Plugin_FileManage_AddControlList In Action_Plugin_FileManage_AddControlList
			If Not IsEmpty(sAction_Plugin_FileManage_AddControlList) Then Call Execute(sAction_Plugin_FileManage_AddControlList)
		Next
		Response.Write "</td></tr>"
		Response.Write vbCrlf 
	Next
	Response.Write "</tbody><tfoot></tfoot></table>"
	Set fold=Nothing

	Set f=Nothing


	FileManage_ExportSiteFileList=True

	Err.Clear


	For Each sAction_Plugin_FileManage_ExportSiteFileList_End In Action_Plugin_FileManage_ExportSiteFileList_End
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteFileList_End) Then Call Execute(sAction_Plugin_FileManage_ExportSiteFileList_End)
	Next
End Function






'*********************************************************
' 目的：    输出编辑文件
'*********************************************************
Function FileManage_ExportSiteFileEdit(tpath,OpenFolderPath,chars)
	For Each sAction_Plugin_FileManage_ExportSiteFileEdit_Begin In Action_Plugin_FileManage_ExportSiteFileEdit_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteFileEdit_Begin) Then Call Execute(sAction_Plugin_FileManage_ExportSiteFileEdit_Begin)
	Next
	

	Dim Del,txaContent
	Dim ct
	Dim cat
	cat=IIf(chars="",FileManage_CheckFileCharset(tpath),UCase(chars))
	ct=TransferHTML(LoadFromFile(unEscape(tpath),cat),"[textarea]")

	'Dim chkg
	'chkg=LCase(BlogPath & unEscape(tpath))
	'if instr(chkg,"global.asa") Then
	'	Response.Write  "<p>当前文件:" & chkg & "</p><p>对不起，为了您的其他程序的安全，您只能修改Z-Blog文件夹内的文件，同时也不允许修改Global.asa和Global.asax。</p><p><a href='main.asp?act=SiteFileMng&path="&Server.URLEncode(OpenFolderPath)&"'>点击这里返回</a></p></div>" :Response.end
	'End If
	If IsEmpty(txaContent) Then txaContent=Null
	
	Response.Write "使用其他编码打开："	
	Response.Write "<input type=""radio"" name=""charset_"" id=""radio_"" value=""UTF-8"" "&IIf(cat="UTF-8","checked=""checked""","")&"/>"
	Response.Write "<label for=""radio_"">UTF-8</label>"
	Response.Write "&nbsp;<input type=""radio"" name=""charset_"" id=""radio2_"" value=""Unicode"" "&IIf(cat="UNICODE","checked=""checked""","")&"/>"
	Response.Write "<label for=""radio2_"">Unicode</label>"
	Response.Write "&nbsp;<input type=""radio"" name=""charset_"" id=""radio3_"" value=""GB2312"" "&IIf(cat="GB2312","checked=""checked""","")&" />"
	Response.Write "<label for=""radio3_"">GB2312</label>"
	Response.Write "<script type=""text/javascript"">$('input[name=charset_]').click(function(){location.search=location.search.replace(/&charset=.+/ig,'')+'&charset='+encodeURIComponent($(this).val())})</script>"
	'Response.Write "<input value=""Go"" type=""submit"" class=""button"">"
	
		
	If Not IsNull(tpath) Then
		Response.Write "<form id=""editdata"" name=""edit"" method=""post"" action=""main.asp?act=SiteFilePst&path="&Server.URLEncode(tpath)&"&OpenFolderPath="&Server.URLEncode(OpenFolderPath)&""">" & vbCrlf
		Response.Write "<p>文件路径及文件名: <!--<a href=""javascript:void(0)"" onclick=""path.readOnly='';this.style.display='none';path.focus()"">修改文件名</a>--><INPUT TYPE=""text"" Value="""&unEscape(tpath)&""" style=""width:100%"" name=""path"" id=""path"" ></p>"
		Response.Write "<p><textarea class=""resizable"" style=""height:300px;width:100%"" name=""txaContent"" id=""txaContent"">"
		Response.Write ct
		Response.Write "</textarea></p>" & vbCrlf

		Response.Write "<hr/>"
		Response.Write "<p><input class=""button"" type=""submit"" value="""&ZC_MSG087&""" id=""btnPost""/>&nbsp;&nbsp;<input class=""button"" type=""button"" value=""返回""  onclick=""location.href='main.asp?act=SiteFileMng&path="&Server.URLEncode(OpenFolderPath)&"'""/>"
		Response.Write "选择文件编码："

		Response.Write "<input type=""radio"" name=""charset"" id=""radio"" value=""UTF-8"" "&IIf(cat="UTF-8","checked=""checked""","")&"/>"
		Response.Write "<label for=""radio"">UTF-8</label>"
		Response.Write "&nbsp;<input type=""radio"" name=""charset"" id=""radio2"" value=""Unicode"" "&IIf(cat="UNICODE","checked=""checked""","")&"/>"
		Response.Write "<label for=""radio2"">Unicode</label>"
		Response.Write "&nbsp;<input type=""radio"" name=""charset"" id=""radio3"" value=""GB2312"" "&IIf(cat="GB2312","checked=""checked""","")&" />"
		Response.Write "<label for=""radio3"">GB2312</label>"
		Response.Write "</p>"

		Response.Write "</form>" & vbCrlf
		If FileManage_OpenCodeMirror Then
    	Response.Write "<script>var editor = CodeMirror.fromTextArea(document.getElementById(""txaContent""), {mode: """
			If CheckRegExp(tpath,".+?html?|.+?xml") Or ct="" Then
				Response.Write 	"text/xml"
			ElseIf CheckRegExp(tpath,".+?js(on)?") Then
				Response.Write  "text/javascript"
			ElseIf CheckRegExp(tpath,".+?css") Then
				Response.Write  "text/css"
			ElseIf CheckRegExp(tpath,".+?asp") Then
				Dim o
				o=Instr(ct,"&lt;script language=""javascript""")
				If o>0 Then
					o=Instr(ct,"runat=""server""")
					If o<400 And o>0 Then
						Response.Write "text/javascript"
					Else
						Response.Write  "application/x-asp"
					End If
				Else
					Response.Write  "application/x-asp"
				End If
				'Response.Write  "application/x-asp"
			End If
			Response.Write """,matchBrackets: true,lineNumbers: true,theme:""monokai"",lineWrapping :true}); </scr"&"ipt>"
		End If
	End If

	If Not FileManage_Return2List___ Then 
		Response.Write "<script type='text/javascript'>$('#editdata').submit(function(){"
		Response.Write "$.post('main.asp?act=SiteFilePst&path="&Server.URLEncode(tpath)&"&OpenFolderPath="&Server.URLEncode(OpenFolderPath)&"',$('#editdata').serialize(),function(data){alert('ok')});return false})</script>"
	End If


	FileManage_ExportSiteFileEdit=True


	For Each sAction_Plugin_FileManage_ExportSiteFileEdit_End In Action_Plugin_FileManage_ExportSiteFileEdit_End
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteFileEdit_End) Then Call Execute(sAction_Plugin_FileManage_ExportSiteFileEdit_End)
	Next
End Function

'*********************************************************
' 目的：    删除文件
'*********************************************************
Function FileManage_DeleteSiteFile(tpath,isFolder)
	For Each sAction_Plugin_FileManage_DeleteSiteFile_Begin In Action_Plugin_FileManage_DeleteSiteFile_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_DeleteSiteFile_Begin) Then Call Execute(sAction_Plugin_FileManage_DeleteSiteFile_Begin)
	Next
	  
	'On Error Resume Next
	Dim SuccessPath
	FileManage_FormatPath tpath
	SuccessPath="main.asp?act=SiteFileMng&path="&Server.URLEncode(Request.QueryString("OpenFolderPath"))
	If isFolder Then
		If FileManage_CheckFolder(tpath)=True Then FileManage_ExportError "不能删除Global.asa和Global.asax和Z-Blog以外的文件夹内的文件",SuccessPath
		If FileManage_ExportInformation("",tpath)<>"" Then
			Call FileManage_ExportError("该文件夹禁止删除！",SuccessPath)
		Else
			If FileManage_FSO.FolderExists(tpath) Then  FileManage_FSO.DeleteFolder(tpath)
		End If
		
	Else
		If FileManage_CheckFile(tpath)=True Then FileManage_ExportError "不能删除Global.asa和Global.asax和Z-Blog以外的文件夹内的文件",SuccessPath
		If FileManage_FSO.FileExists(tpath) Then FileManage_FSO.DeleteFile(tpath)
	End If
	If Err.Number=0 Then
		Call SetBlogHint(True,Empty,Empty)
	Else
		Call FileManage_ExportError("出现错误" & Hex(Err.Number) & "，描述为" & Err.Description & "，操作没有生效",SuccessPath)
	End If
	 	

	For Each sAction_Plugin_FileManage_DeleteSiteFile_End In Action_Plugin_FileManage_DeleteSiteFile_End
		If Not IsEmpty(sAction_Plugin_FileManage_DeleteSiteFile_End) Then Call Execute(sAction_Plugin_FileManage_DeleteSiteFile_End)
	Next
End Function

'*********************************************************
' 目的：    下载文件
'*********************************************************
Function FileManage_DownloadFile(ByVal tpath)
	For Each sAction_Plugin_FileManage_DownloadFile_Begin In Action_Plugin_FileManage_DownloadFile_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_DownloadFile_Begin) Then Call Execute(sAction_Plugin_FileManage_DownloadFile_Begin)
	Next
	
	On Error Resume Next

	FileManage_FormatPath tpath
	
	Dim objGetFile,objADO
	Set objGetFile=FileManage_FSO.getfile(tPath) 
	If FileManage_CheckFile(tpath) Then Response.Write "不能下载Z-Blog以外的文件夹内的文件"
	Response.Clear
	Response.ContentType = "application/octet-stream " 
	Response.AddHeader "Content-Disposition",   "attachment;filename="&objGetFile.name  
	'Response.AddHeader "Content-Length",objGetFile.size  
	Set objADO=Server.CreateObject("ADODB.Stream")
	With objADO
		.Type=adTypeBinary
    	.Mode=adModeReadWrite
    	.Open 
		.Position = objAdo.Size 
    	.LoadFromFile tpath 
		Response.BinaryWrite .Read
		.Close
	End With  
	Response.End 
	'我讨厌打常量。。。。。。
	Set objGetFile=Nothing 
	 
	Set objADO=Nothing

	For Each sAction_Plugin_FileManage_DownloadFile_End In Action_Plugin_FileManage_DownloadFile_End
		If Not IsEmpty(sAction_Plugin_FileManage_DownloadFile_End) Then Call Execute(sAction_Plugin_FileManage_DownloadFile_End)
	Next
End Function

'*********************************************************
' 目的：    重命名文件\文件夹
'*********************************************************
Function FileManage_RenameFile(tpath,newname,isFolder)
	For Each sAction_Plugin_FileManage_RenameFile_Begin In Action_Plugin_FileManage_RenameFile_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_RenameFile_Begin) Then Call Execute(sAction_Plugin_FileManage_RenameFile_Begin)
	Next
	
	On Error Resume Next
	Dim SuccessPath
	FileManage_FormatPath tpath
	SuccessPath="main.asp?act=SiteFileMng&path="&Server.URLEncode(Request.QueryString("OpenFolderPath"))
	If FileManage_CheckFile(tpath)=True Then FileManage_ExportError "不能重命名Global.asa和Global.asax和Z-Blog以外的文件夹内的文件",SuccessPath
	If isFolder Then
		If FileManage_ExportInformation("",tpath)<>"" Then
			Call FileManage_ExportError("该文件夹禁止重命名！",SuccessPath)
		Else
			FileManage_FSO.GetFolder(tpath).name=newname
		End If
	Else
		FileManage_FSO.GetFile(tpath).name=newname
	End If
	If Err.Number=0 Then
		Call SetBlogHint(True,Empty,Empty)
	Else
		Call FileManage_ExportError("出现错误" & Hex(Err.Number) & "，描述为" & Err.Description & "，操作没有生效",SuccessPath)
	End If
	
	Set objGetFile=Nothing 
	 

	For Each sAction_Plugin_FileManage_RenameFile_End In Action_Plugin_FileManage_RenameFile_End
		If Not IsEmpty(sAction_Plugin_FileManage_RenameFile_End) Then Call Execute(sAction_Plugin_FileManage_RenameFile_End)
	Next
End Function


'*********************************************************
' 目的：    输出上传
'*********************************************************
Function FileManage_ExportSiteUpload(path)
	For Each sAction_Plugin_FileManage_ExportSiteUpload_Begin In Action_Plugin_FileManage_ExportSiteUpload_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteUpload_Begin) Then Call Execute(sAction_Plugin_FileManage_ExportSiteUpload_Begin)
	Next
	
	Dim filePath
	Response.Write "<form border=""1"" name=""edit"" id=""edit"" method=""post"" enctype=""multipart/form-data"" action=""main.asp?act=SiteFileUpload"">"
	Response.Write "<p><label for=""path"">保存路径</label><input type=""text"" id=""path"" name=""path"" style=""width:80%"" value="""
	 if instr(path,":")>0 then
	 	filePath=path
	 else
		filePath=BlogPath & path
	 end if
	Response.Write filePath&"""/></p>"
	Response.Write "<p><input type=""file"" id=""edtFileLoad"" name=""edtFileLoad"" size=""20"" style=""width:70%"">"
	Response.Write "<input class=""button"" type=""reset"" value="""& ZC_MSG088 &""" name=""B2"" />"
	Response.Write "<input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" name=""B1"" onclick='' />"
	
	Response.Write "</p></form>"

	For Each sAction_Plugin_FileManage_ExportSiteUpload_End In Action_Plugin_FileManage_ExportSiteUpload_End
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteUpload_End) Then Call Execute(sAction_Plugin_FileManage_ExportSiteUpload_End)
	Next
End Function

'*********************************************************
' 目的：    上传文件
'*********************************************************
Function FileManage_Upload()
	On Error Resume Next
	For Each sAction_Plugin_FileManage_Upload_Begin In Action_Plugin_FileManage_Upload_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_Upload_Begin) Then Call Execute(sAction_Plugin_FileManage_Upload_Begin)
	Next
	Dim objUpload
	Set objUpload=New UpLoadClass
	objUpload.AutoSave=2
	objUpload.Charset="UTF-8"
	objUpload.FileType=""
	objUpload.MaxSize=2^31-1
	objUpload.Open
	Dim tpath,opath,SuccessPath
	tpath=objUpload.Form("path")
	SuccessPath="main.asp?act=SiteFileMng&path="&Server.URLEncode(tpath)
	Dim isOK
	isOK=True
	If FileManage_CheckFile(tpath) Then FileManage_ExportError "不能上传Global.asa和Global.asax，也不能往Z-Blog以外的文件夹上传文件。",SuccessPath
	
	objUpload.SavePath=tpath&"\"
	objUpload.Open
	objUpload.Save "edtFileLoad",1
	If Err.Number=0 And objUpload.Form("edtFileLoad_Err")=0 Then
		Call SetBlogHint(True,Empty,Empty)
	Else
		FileManage_ExportError objUpload.Error2Info("edtFileLoad"),SuccessPath
	End If
	
	Response.Redirect SuccessPath

	For Each sAction_Plugin_FileManage_Upload_End In Action_Plugin_FileManage_Upload_End
		If Not IsEmpty(sAction_Plugin_FileManage_Upload_End) Then Call Execute(sAction_Plugin_FileManage_Upload_End)
	Next
End Function

'*********************************************************
' 目的：    保存文件
'*********************************************************
Function FileManage_PostSiteFile(tpath,OpenFolderPath)
	For Each sAction_Plugin_FileManage_PostSiteFile_Begin In Action_Plugin_FileManage_PostSiteFile_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_PostSiteFile_Begin) Then Call Execute(sAction_Plugin_FileManage_PostSiteFile_Begin)
	Next
	Dim SuccessPath
	SuccessPath="main.asp?act=SiteFileMng&path="

	'On Error Resume Next
	
	FileManage_FormatPath tpath
	If FileManage_FSO.FileExists(tpath) Then
		SuccessPath=SuccessPath&Server.URLEncode(FileManage_FSO.getFile(tpath).ParentFolder)
	Else
		SuccessPath=SuccessPath&Server.URLEncode(OpenFolderPath)
	End If
	If FileManage_CheckFile(tpath)=True Then FileManage_ExportError "不能修改Global.asa和Global.asax和Z-Blog以外的文件夹内的文件",SuccessPath
	Dim txaContent
	txaContent=Request.Form("txaContent")
	If IsEmpty(txaContent) Then txaContent=Null
	If Not IsNull(tpath) Then
		If Not IsNull(txaContent) Then
				Call SaveToFile(tpath,txaContent,Request.Form("charset"),False)
			If Err.Number=0 Then
				Call SetBlogHint(True,Empty,Empty)
				FileManage_PostSiteFile=True
			Else
				FileManage_ExportError "出现错误" & Hex(Err.Number) & "，描述为" & Err.Description & "，操作没有生效。",SuccessPath
			End If
		End IF
	End If
	
	If Not FileManage_Return2List___ Then SuccessPath=Request.ServerVariables("HTTP_REFERER")
	SuccessPath=Replace(SuccessPath,"\","\\")
	Response.Write "<script type=""text/javascript"">location.href="""&SuccessPath&"""</script>"
	Response.End


	For Each sAction_Plugin_FileManage_PostSiteFile_End In Action_Plugin_FileManage_PostSiteFile_End
		If Not IsEmpty(sAction_Plugin_FileManage_PostSiteFile_End) Then Call Execute(sAction_Plugin_FileManage_PostSiteFile_End)
	Next
End Function

'*********************************************************
' 目的：    创建文件夹
'*********************************************************
Function FileManage_CreateFolder(tpath,openpath)
	For Each sAction_Plugin_FileManage_CreateFolder_Begin In Action_Plugin_FileManage_CreateFolder_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_CreateFolder_Begin) Then Call Execute(sAction_Plugin_FileManage_CreateFolder_Begin)
	Next
	Dim SuccessPath
	SuccessPath="main.asp?act=SiteFileMng&path="&Server.UrlEncode(tpath)
	On Error Resume Next
	FileManage_FSO.CreateFolder tpath
	If Err.Number=0 Then
		Call SetBlogHint(True,Empty,Empty)
	Else
		Call FileManage_ExportError("<font color='red'>出现错误" & Hex(Err.Number) & "，描述为" & Err.Description & "，操作没有生效。</font>","main.asp?act=SiteFileMng&path="&Server.URLEncode(openpath))
	End If

	For Each sAction_Plugin_FileManage_CreateFolder_End In Action_Plugin_FileManage_CreateFolder_End
		If Not IsEmpty(sAction_Plugin_FileManage_CreateFolder_End) Then Call Execute(sAction_Plugin_FileManage_CreateFolder_End)
	Next
End Function



'*********************************************************
' 目的：    输出错误
'*********************************************************
Sub FileManage_ExportError(Msg,Url)
	On Error Resume Next
	Response.Clear
	Call SetBlogHint_Custom("<span style='color:red' id='_err'>"&Msg&"</span><script>$('#_err').parent().parent().addClass('hint_red').removeClass('hint_teal');</script>")
	
End Sub

'*********************************************************
' 目的：    检查文件夹是否合法
'*********************************************************
Function FileManage_CheckFolder(folder)
	
	FileManage_CheckFolder=False
	Dim Temp1,Temp2
	If FileManage_FSO.FolderExists(folder)=False Then
		FileManage_CheckFolder=True
	Else
		Temp1=FileManage_FSO.GetFolder(BlogPath).Path
		Temp2=FileManage_FSO.GetFolder(folder).Path
		If Left(Temp2,Len(Temp1))<>Temp1 Then FileManage_CheckFolder=True
	End If 
End Function
Function FileManage_CheckFile(file)
	
	FileManage_CheckFile=False
	
	Dim Temp1,Temp2,Temp3
	'If FileManage_FSO.FileExists(file)=False Then
	'	FileManage_CheckFile=True
	'Else
		Temp1=FileManage_FSO.GetFolder(BlogPath).Path
		If FileManage_FSO.FileExists(file)=True Then
			Temp2=FileManage_FSO.GetFile(file).ParentFolder
			Temp3=LCase(FileManage_FSO.GetFile(file).Name)
			If Left(Temp2,Len(Temp1))<>Temp1 Then FileManage_CheckFile=True
		Else
			Temp3=file
			If Instr(Temp3,Temp1)<=0 Then FileManage_CheckFile=True
		End If
		If CheckRegExp(Temp3,".*?global.asa(x)?") Then FileManage_CheckFile=True
	'End If 
End Function


Sub FileManage_FormatPath(ByRef Path)
	if path<>"" then
		if instr(path,":")>0 then
			path=path
		else
			path=Server.mappath(path)
		end if
	else
		path=BlogPath
	end if

End Sub

Function FileManage_CheckFileCharset(path)
	On Error Resume Next
	Dim binHead
	Dim objStream
	Set objStream=Server.CreateObject("adodb.stream")
	objStream.Type=1
	objStream.mode=3
	objStream.open
	objStream.Position=0
	objStream.LoadFromFile path
	binHead=objstream.read(2)
	If Err.Number=0 Then
		If LenB(binHead)=2 Then
			If AscB(MidB(binHead,1,1))=&HEF And AscB(MidB(binHead,2,1))=&HBB Then
				FileManage_CheckFileCharset="UTF-8"
			ElseIf AscB(MidB(binHead,1,1))=&HFF And AscB(MidB(binHead,2,1))=&HFE Then
				FileManage_CheckFileCharset="UNICODE"
			Else
				FileManage_CheckFileCharset="GB2312"
			End If
		Else
			FileManage_CheckFileCharset="GB2312"
		End If
	Else
		FileManage_CheckFileCharset="UTF-8"
	End If
	Err.Clear
	objStream.Close
	Set objStream=Nothing
End Function

Function FileManage_Setting()
	Response.Write "<form method=""post"" action=""?act=SaveSetting"">"
	Response.Write "<table width=""100%""><tr height=""32""><th width=""300"">配置项</th><th>内容</th></tr>"
	Response.Write "<tr height=""32""><td>打开CodeMirror代码高亮</td><td><input type=""text"" class=""checkbox"""&_
					" name=""OpenCodeMirror"" value="""&FileManage_OpenCodeMirror&"""/></td></tr>"
	Response.Write "<tr height=""32""><td>显示主题名和插件名</td><td><input type=""text"" class=""checkbox"""&_
					" name=""ShowAppsName__"" value="""&FileManage_ShowAppsName__&"""/></td></tr>"
	Response.Write "<tr height=""32""><td>保存文件是否回到文件列表</td><td><input type=""text"" class=""checkbox"""&_
					" name=""Return2List___"" value="""&FileManage_Return2List___&""" style=""width:70%""/></td></tr>"
	Response.Write "<tr height=""32""><td>点击文件管理时打开路径（相对路径）</td><td><input type=""text"" "&_
					" name=""DefaultPath___"" value="""&FileManage_DefaultPath___&""" style=""width:70%""/></td></tr>"
	Response.Write "</table><p>&nbsp;</p>"
	Response.Write "<input type=""submit"" class=""button"" value=""保存""/></form>"
	
End Function

Function FileManage_SaveSetting()
	objConfig.Write "ShowAppsName__",Request.Form("ShowAppsName__")
	objConfig.Write "OpenCodeMirror",Request.Form("OpenCodeMirror")
	objConfig.Write "DefaultPath___",Request.Form("DefaultPath___")
	objConfig.Write "Return2List___",Request.Form("Return2List___")
	objConfig.Save
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "?act=Setting"
End Function
%>