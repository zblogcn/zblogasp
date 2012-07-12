<!-- #include file="include_plugin.asp"-->
<%
'*********************************************************
' 目的：    格式化文件大小
'*********************************************************
Function FileManage_GetSize(FileSize)
	For Each sAction_Plugin_FileManage_GetSize_Begin in Action_Plugin_FileManage_GetSize_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_GetSize_Begin) Then Call Execute(sAction_Plugin_FileManage_GetSize_Begin)
	Next
	
	Dim b,m
	b=filesize:m="B"
	if b>1024 then b=b/1024:m="K"
	if b>1024 then b=b/1024:m="M"
	if b>1024 then b=b/1024:m="G"
	b=formatnumber(b,2)
	FileManage_GetSize=b&m

	For Each sAction_Plugin_FileManage_GetSize_End in Action_Plugin_FileManage_GetSize_End
		If Not IsEmpty(sAction_Plugin_FileManage_GetSize_End) Then Call Execute(sAction_Plugin_FileManage_GetSize_End)
	Next
End Function

'*********************************************************
' 目的：    得到文件图标
'*********************************************************
Function FileManage_GetTypeIco(FileName)
	For Each sAction_Plugin_FileManage_GetTypeIco_Begin in Action_Plugin_FileManage_GetTypeIco_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_GetTypeIco_Begin) Then Call Execute(sAction_Plugin_FileManage_GetTypeIco_Begin)
	Next
	
	Dim aryFn
	aryFn=Split(FileName,".")
	Dim strType
	strType=LCase(aryFn(Ubound(aryFn)))
	Dim ImgTag,Tag
	ImgTag="<img width=""11"" height=""11"" src=""ico\{tag}.png""/>"
	Select Case strType
		Case "txt","doc","ppt","xls","rtf","ini","inf","sql","log" Tag="txt"
		Case "mp3","wma","wav","ogg" Tag="msc"
		Case "mpg","mpeg","avi","rm","rmvb","vob","dat","mp4","3gp","flv","swf","mkv","mov" Tag="mov"
		Case "exe","com" Tag="exe"
		Case "dll","ocx","sys","db" Tag="dll"
		Case "bat","cmd" Tag="bat"
		Case "asp","php","jsp","js","css","inc","asa","asax","aspx","mhtml","shtml","py"  Tag="code"
		Case "jpg","jpeg","gif","bmp","png","tiff" Tag="img"
		Case "htm","html","xml"  Tag="htm"
		Case "rar","zip","7z","gz"  Tag="rar"
		Case "mdb" Tag="mdb"

		Case Else  Tag="no"
	End Select
	FileManage_GetTypeIco=Replace(ImgTag,"{tag}",tag)

	For Each sAction_Plugin_FileManage_GetTypeIco_End in Action_Plugin_FileManage_GetTypeIco_End
		If Not IsEmpty(sAction_Plugin_FileManage_GetTypeIco_End) Then Call Execute(sAction_Plugin_FileManage_GetTypeIco_End)
	Next
End Function
'*********************************************************
' 目的：    输出注释
'*********************************************************

Function FileManage_ExportInformation(foldername,path)
	For Each sAction_Plugin_FileManage_ExportInformation_Begin in Action_Plugin_FileManage_ExportInformation_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_ExportInformation_Begin) Then Call Execute(sAction_Plugin_FileManage_ExportInformation_Begin)
	Next

	Dim z,k,l,n
	n=""
	z=LCase(foldername)
	k=LCase(path)
	l=lcase(blogpath)
	if k=l then
		select case z
			case "zb_system" n="Z-Blog系统核心文件"
			case "zb_users" n="Z-Blog用户配置文件夹"
			case "zb_install" n="Z-Blog安装文件夹"
			case lcase(ZC_STATIC_DIRECTORY) n="静态文件存放文件夹"
			case "catalog.asp" n="文章列表"
			case "default.asp" n="首页"
			case "feed.asp" n="RSS订阅"
			case "search.asp" n="搜索"
			case "tags.asp" n="Tags列表"
			case "wap.asp" n="Wap"
		end select
	elseif k=l & "\zb_system" then
		select case z
			case "admin" n="Z-Blog系统管理文件"
			case "css" n="Z-Blog后台CSS存放文件夹"
			case "function" n="Z-Blog核心文件"
			case "image" n="Z-Blog后台图片存放文件夹"
			case "script" n="Z-Blog脚本存放文件夹"
			case "wap" n="Z-Blog Wap存放文件夹"
			case "xml-rpc" n="Z-Blog Xml-Rpc存放文件夹"
		end select
	elseif k=l & "\zb_users" then
		select case z
			case "cache" n="Z-Blog缓存文件夹"
			case "data" n="Z-Blog数据库存放文件夹"
			case "include" n="Z-Blog引用文件夹"
			case "language" n="Z-Blog Language Pack"
			case "plugin" n="Z-Blog 插件文件夹"
			case "theme" n="Z-Blog 主题存放文件夹"
			case lcase(ZC_UPLOAD_DIRECTORY) n="上传文件存放文件夹"
			case "c_custom.asp" n="用户配置文件"
			case "c_option.asp" n="网站设置文件"
			case "c_option_wap.asp" n="Wap设置文件"
		end select
	elseif k=l &  "\zb_users\include" then
		select case z
			case "link.asp" n="友链"
			case "favorite.asp" n="收藏"
			case "navbar.asp" n="导航栏"
			case "misc.asp" n="图标汇总"
		end select
	elseif k=l &  "\zb_users\data" then
			if CheckRegExp(z,".+?mdb|.+?asp") then n="可能是Z-Blog数据库"
	elseif k=l & "\zb_users\theme\" & lcase(ZC_BLOG_THEME) then 
		select case z
			case "include" n="引用"
			case "plugin" n="主题自带插件"
			case "source" n="主题CSS"
			case "style" n="主题CSS"
			case lcase(ZC_TEMPLATE_DIRECTORY) n="主题模板"
		end select
	elseif k=l & "\zb_users\theme\"&lcase(zc_blog_theme)&"\"&lcase(ZC_TEMPLATE_DIRECTORY) then
		z=split(z,".")(0)
		select case z
			case "b_article-istop" n= "首页置顶文章模板"
			case "b_article-multi" n= "首页摘要文章模板"
			case "b_article-single" n= "日志页文章模板"
			case "b_article-guestbook" n= "留言页正文模板"
			case "b_article_comment" n= "每条评论内容显示模板"
			case "b_article_commentpost-verify" n= "评论验证码显示样式"
			case "b_article_commentpost" n= "评论发表框模板"
			case "b_article_mutuality" n= "每条相关文章显示模板"
			case "b_article_nvabar_l" n= "“上一篇”日志链接"
			case "b_article_nvabar_r" n= "“下一篇”日志链接"
			case "b_article_tag" n="Tag显示样式"
			case "b_pagebar" n="分页条模板"
			case "catalog" n="分类页整页模板"
			case "default" n="首页整页模板"
			case "search" n="搜索页整页模板"
			case "single" n="日志页整页模板"
			case "tags" n="标签页整页模板"
			case "guestbook" n="留言页整页模板"
		end select
	elseif k=l &"\zb_system\admin" then
		select case z
			case "admin.asp" n="管理页"
			case "admin_default.asp" n="主面板"
			case "admin_left.asp" n="左侧面板"
			case "admin_top.asp" n="后台头文件"
			case "c_autosaverjs.asp" n="自动保存"
			case "c_updateinfo.asp" n="得到最新消息"
			case "edit_catalog.asp" n="编辑分类页"
			case "edit_comment.asp" n="编辑评论页"
			case "edit_link.asp" n="链接管理页"
			case "edit_setting.asp" n="网站设置页"
			case "edit_tag.asp" n="Tag修改页"
			case "edit_ueditor.asp" n="新建文章页"
			case "edit_user.asp" n="用户编辑页"
			case "ueditor" n="Ueditor主文件"
		end select
	elseif k=l & "\zb_system\admin\ueditor" then
		select case z
			case "asp" n="uEditor ASP后台"
			case "dialogs" n="uEditor 对话框"
			case "themes" n="uEditor 主题"
			case "third-party" n="第三方组件"
			case "editor_all_min.js" n="uEditor"
			case "editor_config.asp" n="uEditor配置"
		end select
	elseif k=l & "\zb_system\admin\ueditor\asp" then
		select case z
			case "fileup.asp" n="文件上传"
			case "getcontent.asp" n="得到内容"
			case "getmovie.asp" n="视频搜索"
			case "getremoteimage.asp" n="下载远程图片"
			case "imagemanager.asp" n="图片管理"
			case "picup.asp" n="图片上传"
			case "up_inc.asp" n="风声无组件上传"
		end select
	elseif k=l & "\zb_system\function" then
		select case z
			case "c_error.asp" n="Z-Blog错误处理"
			case "c_function.asp" n="Z-Blog一般函数"
			case "c_html_js.asp" n="访问计数等JS调用"
			case "c_html_js_add.asp" n="动态JS调用文件"
			case "c_system_base.asp" n="Z-Blog基础"
			case "c_system_event.asp" n="Z-Blog事件"
			case "c_system_lib.asp" n="Z-Blog 数据库访问类"
			case "c_system_manage.asp" n="Z-Blog 后台管理文件"
			case "c_system_plugin.asp" n="Z-Blog 插件支持文件"
			case "c_system_wap.asp" n="Z-Blog Wap支持文件"
			case "c_urlredirect.asp" n="Z-Blog 加密Url跳转页"
			case "c_validcode.asp" n="Z-Blog验证码"
		end select
	elseif k=l & "\zb_system\wap" then
		select case z
			case "default.asp" n="Wap首页"
			case "index.asp" n="Wap首页"
			case "style" n="WapCSS"
			case "wap_article-multi.html" n="Wap模板-文章"
			case "wap_article_comment.html" n="Wap模板-评论"
			case "wap_single.html" n="Wap模板-文章页或列表页"
		end select
	end if
	For Each sAction_Plugin_FileManage_ExportInformation_End in Action_Plugin_FileManage_ExportInformation_End
		If Not IsEmpty(sAction_Plugin_FileManage_ExportInformation_End) Then Call Execute(sAction_Plugin_FileManage_ExportInformation_End)
	Next
	FileManage_ExportInformation=n
End Function
'*********************************************************
' 目的：    输出文件列表
'*********************************************************
Function FileManage_ExportSiteFileList(path,opath)
	For Each sAction_Plugin_FileManage_ExportSiteFileList_Begin in Action_Plugin_FileManage_ExportSiteFileList_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteFileList_Begin) Then Call Execute(sAction_Plugin_FileManage_ExportSiteFileList_Begin)
	Next
	

	'On Error Resume Next
	dim f,fold,item,fpath,jpath
	If opath<>"" Then path=opath

	  if path<>"" then
		 if instr(path,":")>0 then
		 path=path
		 else
		 path=server.mappath(path)
		 end if
	  else
	  path=BlogPath
	  end if
	dim backfolder
	backfolder=split(path,"\")
	redim preserve backfolder(ubound(backfolder)-1)
	backfolder=join(backfolder,"\")
	  if FileManage_CheckFolder(path) Then Response.Write  "<p>当前路径:" & path & "</p><p>对不起，为了您的其他程序的安全，您只能修改Z-Blog文件夹内的文件，同时也不允许修改Global.asa和Global.asax。</p><p><a href='main.asp?act=SiteFileMng&path="&Server.URLEncode(BlogPath)&"'>点击这里返回</a></p></div>" :Response.end
	set f=server.createobject("scripting.filesystemobject")


	response.write "<p>"&ZC_MSG240&":"&path&"</p>"
	set fold=f.getfolder(path)

	Response.write"<table width=""100%"" border=""0"">"
	Response.write "<tr><td colspan=""5""><a href='main.asp?act=SiteFileMng&path="&Server.URLEncode(backfolder)&"' title='"&ZC_MSG239&"'><img src=""ico\up.png""/></a>"
	Response.write "&nbsp;&nbsp;<a href='javascript:void(0)' onclick='window.open(""main.asp?act=SiteFileUploadShow&path="&Server.URLEncode(fpath)&"&opath="& Server.URLEncode(path) &""",""Detail"",""Scrollbars=no,Toolbar=no,Location=no,Direction=no,Resizeable=no,height=165px,width=780px"")' title=""上传""><img src=""ico\upload.png""/></a>"
	Response.Write "&nbsp;&nbsp;<a href='main.asp?act=SiteCreateFolder' onmousedown=""var str=prompt('请输入文件夹名');if(str!=null){this.href+='&path='+encodeURIComponent('"&Replace(Replace(path,"\","\\"),"""","\""")&"'+'\\'+str);this.click()}else{return false}"" title='新建文件夹'><img src='ico\cfolder.png'/></a><span style=""float:right""><a href=""main.asp?act=Help"" title=""帮助""><img src=""ico\hlp.png""/></a></span>"
	Response.Write "&nbsp;&nbsp;<a href=""main.asp?act=SiteFileEdt&path="&Server.URLEncode(path) &""" title=""创建文件""><img src=""ico\newfile.png""/></a></td></tr>"
	Response.write "<tr><td width=""20%"">文件名</td><td width=""10%"">修改时间</td><td width=""8%"">大小</td><td width=""10%"">注释</td><td>操作</td></tr>"
	for each item in fold.subfolders
		jpath=replace(path,"\","\\")
		Response.write "<tr height=18><td><img width=""11"" height=""11""src='ico/fld.png' />&nbsp;<a href='main.asp?act=SiteFileMng&path="&Server.URLEncode(path&"\"&item.name)&"&opath='>"&item.name&"</a>"
		Response.write"</td><td>"&item.datelastmodified&"</td><td></td><td>"&FileManage_ExportInformation(item.name,path)&"</td><td></td></tr>"
	next
	for each item in fold.files
	fpath=replace(path&"/"&item.name,BlogPath,"")
	fpath=replace(fpath,"\","/")
	Response.write "<tr><td>"&FileManage_GetTypeIco(item.name)&"&nbsp;<a href=""javascript:;"" title='"&ZC_MSG261&":"&item.datelastmodified&";"&ZC_MSG238&":"&clng(item.size/1024)&"k'>"&item.name&"</a></td><td>"&item.datelastmodified&"</td><td>"&FileManage_GetSize(item.size)&"</td><td>"&FileManage_ExportInformation(item.name,path)&"</td><td>"
	Response.write"<a href=""main.asp?act=SiteFileEdt&path="&Server.URLEncode(fpath)&"&opath="& Server.URLEncode(path) &""" title=""["&ZC_MSG078&"]""><img src=""ico\edit.png"" width=""11"" height=""11""/></a>"
	Response.Write "&nbsp;&nbsp;<a href=""main.asp?act=SiteFileDownload&path="&Server.URLEncode(fpath)&"&opath="& Server.URLEncode(path) &""" target=""_blank"" title=""[下载]""><img src=""ico\download.png"" width=""11"" height=""11""/></a>"
	Response.Write "&nbsp;&nbsp;<a href=""main.asp?act=SiteFileDel&path="&Server.URLEncode(fpath)&"&opath="& Server.URLEncode(path) &""" onclick='return window.confirm("""&ZC_MSG058&""");' title=""["&ZC_MSG063&"]""><img src=""ico\del.png"" width=""11"" height=""11""/></a>"
	Response.Write "&nbsp;&nbsp;<a href=""main.asp?act=SiteFileRename&path="&Server.URLEncode(fpath)&"&opath="& Server.URLEncode(path) &""" onmousedown='var str=prompt(""请输入新文件名"");if(str!=null){this.href+=""&newfilename=""+encodeURIComponent(str);this.click()}else{return false}' title=""[重命名]""><img src=""ico\rename.png"" width=""11"" height=""11""/></a></td></tr>"

	next
	response.write"</table>"
	set fold=nothing

	set f=Nothing


	FileManage_ExportSiteFileList=True

	Err.Clear


	For Each sAction_Plugin_FileManage_ExportSiteFileList_End in Action_Plugin_FileManage_ExportSiteFileList_End
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteFileList_End) Then Call Execute(sAction_Plugin_FileManage_ExportSiteFileList_End)
	Next
End Function






'*********************************************************
' 目的：    输出编辑文件
'*********************************************************
Function FileManage_ExportSiteFileEdit(tpath,opath)
	For Each sAction_Plugin_FileManage_ExportSiteFileEdit_Begin in Action_Plugin_FileManage_ExportSiteFileEdit_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteFileEdit_Begin) Then Call Execute(sAction_Plugin_FileManage_ExportSiteFileEdit_Begin)
	Next
	

	Dim Del,txaContent


	'dim chkg
	'chkg=lcase(BlogPath & unEscape(tpath))
	'if instr(chkg,"global.asa") Then
	'	Response.Write  "<p>当前文件:" & chkg & "</p><p>对不起，为了您的其他程序的安全，您只能修改Z-Blog文件夹内的文件，同时也不允许修改Global.asa和Global.asax。</p><p><a href='main.asp?act=SiteFileMng&path="&Server.URLEncode(oPath)&"'>点击这里返回</a></p></div>" :Response.end
	'End If
	If IsEmpty(txaContent) Then txaContent=Null

		
	If Not IsNull(tpath) Then

		Response.Write "<form id=""edit"" name=""edit"" method=""post"" action=""main.asp?act=SiteFilePst&path="&Server.URLEncode(tpath)&"&opath="&Server.URLEncode(opath)&""">" & vbCrlf
		
		Response.Write "<p><br/>文件路径及文件名: <!--<a href=""javascript:void(0)"" onclick=""path.readOnly='';this.style.display='none';path.focus()"">修改文件名</a>--><INPUT TYPE=""text"" Value="""&unEscape(tpath)&""" style=""width:100%"" name=""path"" id=""path"" ></p>"
		Response.Write "<p><textarea class=""resizable"" style=""height:300px;width:100%"" name=""txaContent"" id=""txaContent"">"&TransferHTML(LoadFromFile(BlogPath & unEscape(tpath),"utf-8"),"[textarea]")&"</textarea></p>" & vbCrlf
		Response.Write "<hr/>"
		Response.Write "<p><input class=""button"" type=""submit"" value="""&ZC_MSG087&""" id=""btnPost""/><input class=""button"" type=""button"" value=""撤销修改，返回""  onclick=""history.go(-1)""/></p>" & vbCrlf
		Response.Write "</form>" & vbCrlf
    	Response.Write "<script>var editor = CodeMirror.fromTextArea(document.getElementById(""txaContent""), {mode: {"
		If CheckRegExp(tpath,".+?html?|.+?xml") Then
			Response.Write 	"name: ""xml"","
		ElseIf CheckRegExp(tpath,".+?js(on)?") Then
			Response.Write  "name: ""javascript"","
		ElseIf CheckRegExp(tpath,".+?css") Then
			Response.Write  "name: ""css"","
		End If
		Response.write " alignCDATA: true},lineNumbers: true}); </script>"
	End If


	FileManage_ExportSiteFileEdit=True


	For Each sAction_Plugin_FileManage_ExportSiteFileEdit_End in Action_Plugin_FileManage_ExportSiteFileEdit_End
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteFileEdit_End) Then Call Execute(sAction_Plugin_FileManage_ExportSiteFileEdit_End)
	Next
End Function

'*********************************************************
' 目的：    删除文件
'*********************************************************
Function FileManage_DeleteSiteFile(tpath)
	For Each sAction_Plugin_FileManage_DeleteSiteFile_Begin in Action_Plugin_FileManage_DeleteSiteFile_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_DeleteSiteFile_Begin) Then Call Execute(sAction_Plugin_FileManage_DeleteSiteFile_Begin)
	Next
	
	On Error Resume Next
	If DelSiteFile(Request.QueryString("path")) Then
		Call SetBlogHint(True,True,Empty)
	Else
		Call SetBlogHint_Custom("<font color='red'>出现错误" & Hex(Err.Number) & "，描述为" & Err.Description & "，操作没有生效。</font>")
	End If
	Response.Write "<script type=""text/javascript"">location.href=""main.asp?act=SiteFileMng" & "&path=" & Replace(Request.QueryString("opath"),"\","\\")&"""</script>"

	For Each sAction_Plugin_FileManage_DeleteSiteFile_End in Action_Plugin_FileManage_DeleteSiteFile_End
		If Not IsEmpty(sAction_Plugin_FileManage_DeleteSiteFile_End) Then Call Execute(sAction_Plugin_FileManage_DeleteSiteFile_End)
	Next
End Function

'*********************************************************
' 目的：    下载文件
'*********************************************************
Function FileManage_DownloadFile(ByVal tpath)
	For Each sAction_Plugin_FileManage_DownloadFile_Begin in Action_Plugin_FileManage_DownloadFile_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_DownloadFile_Begin) Then Call Execute(sAction_Plugin_FileManage_DownloadFile_Begin)
	Next
	
	On Error Resume Next
	Dim filePath,isOK,i,fxxxPath
	fxxxPath=Replace(LCase(BlogPath),"zb_users\plugin\filemanage\..\..\..\","")
	 if instr(tpath,":")>0 then
	 	filePath=tpath
	 else
		filePath=fxxxPath & tpath
	 end if
	 filepath=replace(filepath,"\./","\")
	Dim objFSO,objGetFile,objADO
	Set objFSO=Server.CreateObject("Scripting.FileSystemObject") 
	Set objGetFile=objFSO.getfile(FilePath) 
	isOK=True
	If Instr(LCase(objGetFile.Path),fxxxPath)=0 Then isOK=False
	If isOK=False Then Response.Write "<script>alert('不能下载Z-Blog以外的文件夹内的文件');window.close()</script>":Response.End
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
    	.LoadFromFile FilePath 
		Response.BinaryWrite .Read
		.Close
	End With  
	'我讨厌打常量。。。。。。
	Set objGetFile=Nothing 
	Set objFSO=Nothing 
	Set objADO=Nothing

	For Each sAction_Plugin_FileManage_DownloadFile_End in Action_Plugin_FileManage_DownloadFile_End
		If Not IsEmpty(sAction_Plugin_FileManage_DownloadFile_End) Then Call Execute(sAction_Plugin_FileManage_DownloadFile_End)
	Next
End Function

'*********************************************************
' 目的：    重命名文件\文件夹
'*********************************************************
Function FileManage_RenameFile(tpath,newname)
	For Each sAction_Plugin_FileManage_RenameFile_Begin in Action_Plugin_FileManage_RenameFile_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_RenameFile_Begin) Then Call Execute(sAction_Plugin_FileManage_RenameFile_Begin)
	Next
	  
	On Error Resume Next
	Dim filePath,isOK,i,fxxxPath
	fxxxPath=Replace(LCase(BlogPath),"zb_users\plugin\filemanage\..\..\..\","")
	 if instr(tpath,":")>0 then
	 	filePath=tpath
	 else
		filePath=fxxxPath & tpath
	 end if
	 filepath=replace(replace(filepath,"\/","\"),"\\","\")
	Dim objFSO,objGetFile,objADO
	Set objFSO=Server.CreateObject("Scripting.FileSystemObject") 
	Set objGetFile=objFSO.getfile(FilePath) 
	isOK=True
	If left(lcase(objGetFile.name),10)="global.asa" Then isOK=False
	If Instr(LCase(objGetFile.Path),fxxxPath)=0 Then isOK=False
	If isOK=False Then Response.Write "<script>alert('不能重命名Global.asa和Global.asax和Z-Blog以外的文件夹内的文件');window.close()</script>":Response.End
	objGetFile.name=newname
	If Err.Number=0 Then
		Call SetBlogHint(True,True,Empty)
	Else
		Call SetBlogHint_Custom("<font color='red'>出现错误" & Hex(Err.Number) & "，描述为" & Err.Description & "，操作没有生效。</font>")
	End If
	
	Response.Write "<script type=""text/javascript"">location.href=""main.asp?act=SiteFileMng" & "&path=" & Replace(Request.QueryString("opath"),"\","\\")&"""</script>"
	Set objGetFile=Nothing 
	Set objFSO=Nothing 

	For Each sAction_Plugin_FileManage_RenameFile_End in Action_Plugin_FileManage_RenameFile_End
		If Not IsEmpty(sAction_Plugin_FileManage_RenameFile_End) Then Call Execute(sAction_Plugin_FileManage_RenameFile_End)
	Next
End Function

'*********************************************************
' 目的：    检查文件夹是否合法
'*********************************************************
Function FileManage_CheckFolder(folder)
	For Each sAction_Plugin_FileManage_CheckFolder_Begin in Action_Plugin_FileManage_CheckFolder_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_CheckFolder_Begin) Then Call Execute(sAction_Plugin_FileManage_CheckFolder_Begin)
	Next
	
	FileManage_CheckFolder=False
	dim sptzsx,xhf,t1,t2,t3
	  sptzsx=split(blogpath,"\")
	  for xhf=0 to ubound(sptzsx)
	  	if sptzsx(xhf)=".." then t2=t2+1 else t1=t1+1
	  next
	t3=t1-t2:t2=0:t1=0
  	  sptzsx=split(folder,"\")
	  for xhf=0 to ubound(sptzsx)
	  	if sptzsx(xhf)=".." then t2=t2+1 else t1=t1+1
	  next
	  if t1-t2<t3 Then FileManage_CheckFolder=True

	For Each sAction_Plugin_FileManage_CheckFolder_End in Action_Plugin_FileManage_CheckFolder_End
		If Not IsEmpty(sAction_Plugin_FileManage_CheckFolder_End) Then Call Execute(sAction_Plugin_FileManage_CheckFolder_End)
	Next
End Function

'*********************************************************
' 目的：    输出上传
'*********************************************************
Function FileManage_ExportSiteUpload(path)
	For Each sAction_Plugin_FileManage_ExportSiteUpload_Begin in Action_Plugin_FileManage_ExportSiteUpload_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteUpload_Begin) Then Call Execute(sAction_Plugin_FileManage_ExportSiteUpload_Begin)
	Next
	
	dim filePath
	Response.Write "<form border=""1"" name=""edit"" id=""edit"" method=""post"" enctype=""multipart/form-data"" action=""main.asp?act=SiteFileUpload"">"
	Response.Write "<p><label for=""path"">请输入保存路径</label><input type=""text"" id=""path"" name=""path"" style=""width:80%"" value="""
	 if instr(path,":")>0 then
	 	filePath=path
	 else
		filePath=BlogPath & path
	 end if
	Response.Write filePath&"""/></p>"
	Response.Write "<p><input type=""file"" id=""edtFileLoad"" name=""edtFileLoad"" size=""20"">  <input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" name=""B1"" onclick='' /> <input class=""button"" type=""reset"" value="""& ZC_MSG088 &""" name=""B2"" />"


	For Each sAction_Plugin_FileManage_ExportSiteUpload_End in Action_Plugin_FileManage_ExportSiteUpload_End
		If Not IsEmpty(sAction_Plugin_FileManage_ExportSiteUpload_End) Then Call Execute(sAction_Plugin_FileManage_ExportSiteUpload_End)
	Next
End Function

'*********************************************************
' 目的：    上传文件
'*********************************************************
Function FileManage_Upload()
	On Error Resume Next
	Dim objUpload
	Set objUpload=New FileManage_UpLoadClass
	objUpload.AutoSave=2
	objUpload.Charset="UTF-8"
	objUpload.FileType=""
	objUpload.open
	Dim tpath
	tpath=objUpload.Form("path")
	For Each sAction_Plugin_FileManage_Upload_Begin in Action_Plugin_FileManage_Upload_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_Upload_Begin) Then Call Execute(sAction_Plugin_FileManage_Upload_Begin)
	Next
	Dim isOK
	isOK=True
	If FileManage_CheckFolder(tpath) Then isOK=False
'	If objUpload.Form("edtFileLoad")="" Then isOK=False
	If Left(LCase(objUpload.Form("edtFileLoad")),10)="global.asa" Then isOK=False
	If isOK=False Then Response.Write "<script>alert('不能上传Global.asa和Global.asax，也不能往Z-Blog以外的文件夹上传文件。同时上传时最大文件大小不能超过200K，否则可能会被IIS限制。');window.close()</script>":Response.End
	objUpload.SavePath=tpath
	objUpload.open
	objUpload.save "edtFileLoad",1
	If Err.Number=0 Then
		Call SetBlogHint(True,True,Empty)
	Else
		Call SetBlogHint_Custom("<font color='red'>出现错误" & Hex(Err.Number) & "，描述为" & Err.Description & "，操作没有生效。</font>")	End If
	
	Response.Write "<script>opener.location.reload();window.close();</script>"

	For Each sAction_Plugin_FileManage_Upload_End in Action_Plugin_FileManage_Upload_End
		If Not IsEmpty(sAction_Plugin_FileManage_Upload_End) Then Call Execute(sAction_Plugin_FileManage_Upload_End)
	Next
End Function

'*********************************************************
' 目的：    保存文件
'*********************************************************
Function FileManage_PostSiteFile(tpath)
	For Each sAction_Plugin_FileManage_PostSiteFile_Begin in Action_Plugin_FileManage_PostSiteFile_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_PostSiteFile_Begin) Then Call Execute(sAction_Plugin_FileManage_PostSiteFile_Begin)
	Next
	
	'On Error Resume Next
	Dim filePath,isOK,i,fxxxPath
	fxxxPath=Replace(LCase(BlogPath),"zb_users\plugin\filemanage\..\..\..\","")
	 if instr(tpath,":")>0 then
	 	filePath=tpath
	 else
		filePath=fxxxPath & tpath
	 end if

	 filepath=replace(replace(filepath,"\/","\"),"\\","\")
	Dim objFSO,objGetFile,objADO
	Set objFSO=Server.CreateObject("Scripting.FileSystemObject") 
	IsOK=True
	If objFSO.FileExists(FilePath) Then
			Set objGetFile=objFSO.getfile(FilePath) 
			If left(lcase(objGetFile.name),10)="global.asa" Then isOK=False
			If Instr(LCase(objGetFile.Path),fxxxPath)=0 Then isOK=False
		Else
			If Instr(lcase(FilePath),"global.asa")>0 Then isOK=False
	End If
	If isOK=False Then Response.Write "<script>alert('不能修改Global.asa和Global.asax和Z-Blog以外的文件夹内的文件');history.go(-1)</script>":Response.End
	Dim txaContent
	txaContent=Request.Form("txaContent")
	If IsEmpty(txaContent) Then txaContent=Null
	If Not IsNull(tpath) Then
		If Not IsNull(txaContent) Then
				Call SaveToFile(FilePath,txaContent,"utf-8",False)
			If Err.Number=0 Then
				Call SetBlogHint(True,True,Empty)
			Else
				Call SetBlogHint_Custom("<font color='red'>出现错误" & Hex(Err.Number) & "，描述为" & Err.Description & "，操作没有生效。</font>")			End If
			FileManage_PostSiteFile=True
		End IF
	End If
	Response.Write "<script type=""text/javascript"">location.href=""main.asp?act=SiteFileMng" & "&path=" & Replace(Request.QueryString("opath"),"\","\\")&"""</script>"


	For Each sAction_Plugin_FileManage_PostSiteFile_End in Action_Plugin_FileManage_PostSiteFile_End
		If Not IsEmpty(sAction_Plugin_FileManage_PostSiteFile_End) Then Call Execute(sAction_Plugin_FileManage_PostSiteFile_End)
	Next
End Function

'*********************************************************
' 目的：    创建文件夹
'*********************************************************
Function FileManage_CreateFolder(tpath)
	For Each sAction_Plugin_FileManage_CreateFolder_Begin in Action_Plugin_FileManage_CreateFolder_Begin
		If Not IsEmpty(sAction_Plugin_FileManage_CreateFolder_Begin) Then Call Execute(sAction_Plugin_FileManage_CreateFolder_Begin)
	Next

	On Error Resume Next
	Dim filePath,isOK,i,fxxxPath
	 if instr(tpath,":")>0 then
	 	filePath=tpath
	 else
		filePath=fxxxPath & tpath
	 end if
	 '创建文件夹时不对是否为zblog之外文件夹判断
	 filepath=replace(replace(filepath,"\/","\"),"\\","\")
	Dim objFSO,objGetFile,objADO
	Set objFSO=Server.CreateObject("Scripting.FileSystemObject") 
	objFSO.CreateFolder tpath
	If Err.Number=0 Then
		Call SetBlogHint(True,True,Empty)
	Else
		Call SetBlogHint_Custom("<font color='red'>出现错误" & Hex(Err.Number) & "，描述为" & Err.Description & "，操作没有生效。</font>")	End If
	Set objFSO=nothing
	Response.Write "<script type=""text/javascript"">location.href=""main.asp?act=SiteFileMng" & "&path=" & Replace(tpath,"\","\\")&"""</script>"

	For Each sAction_Plugin_FileManage_CreateFolder_End in Action_Plugin_FileManage_CreateFolder_End
		If Not IsEmpty(sAction_Plugin_FileManage_CreateFolder_End) Then Call Execute(sAction_Plugin_FileManage_CreateFolder_End)
	Next
End Function



Function FileManage_Help
	%>
    	<style>
		ol {line-height:220%;}
		ol li {margin:0 0 0 -18px;text-decoration: none;}
		b {color:Navy;font-weight:Normal;text-decoration: underline;}
		p {line-height:160%;}
	</style>
    <p>您正在使用的插件，是由ZSXSOFT制作的强化Z-Blog文件管理的插件。
    </p>
<ol>
  <li>插件拥有功能：上传、下载、重命名、删除、编辑、新增文件、新建文件夹；</li>
  <li>由于用户误操作而对网站造成任何损害（包括但不限于Z-Blog无法打开、数据被破坏等），插件原作者已经尽到提醒责任，没有解决问题的义务。</li>
  <li>由于批量删除文件（夹）、重命名文件夹过于危险，所以没有开放；</li>
  <li>为了保证您的服务器安全，插件有如下限制：
    <ol>
      <li>不允许修改Global.asa和Global.asax以防止全站挂马</li>
      <li>不允许任何对Z-Blog以外的文件（夹）操作。</li>
    </ol>
  </li>
  <li>Ico文件夹内部分图标来自Microsoft Corporation、Adobe Software、RARLAB。</li>
  <li>插件接口以及注意事项如下：
    <ol>
      <li>Action类接口（使用方法请参见<a href="http://wiki.rainbowsoft.org/doku.php?id=plugin:api:action">Z-Wik</a>i）
        <ol>
          <li>            Action_Plugin_FileManage_Initialize         当文件管理页面加载时被触发</li>
          <li>Action_Plugin_FileManage_Terminate 当文件管理页面加载完毕后被触发</li>
          <li>Action_Plugin_FileManage_ExportSiteFileList_Begin 当加载文件列表时被触发</li>
          <li>Action_Plugin_FileManage_ExportSiteFileList_End 当加载文件列表结束后被触发</li>
          <li>Action_Plugin_FileManage_GetTypeIco_Begin 当得到文件图标时被触发</li>
          <li>Action_Plugin_FileManage_GetTypeIco_End 当文件图标获取完毕后被触发</li>
          <li>Action_Plugin_FileManage_ExportSiteFileEdit_Begin 当加载编辑器时被触发</li>
          <li>Action_Plugin_FileManage_ExportSiteFileEdit_End 当加载编辑器完毕后被触发</li>
          <li>Action_Plugin_FileManage_DeleteSiteFile_Begin 当删除文件时被触发</li>
          <li>Action_Plugin_FileManage_DeleteSiteFile_End 当删除完毕后被触发</li>
          <li>Action_Plugin_FileManage_DownloadFile_Begin 当下载文件时被触发</li>
          <li>Action_Plugin_FileManage_DownloadFile_End 当下载文件后被触发</li>
          <li>Action_Plugin_FileManage_RenameFile_Begin 当改名时被触发</li>
          <li>Action_Plugin_FileManage_RenameFile_End 当改名后被触发</li>
          <li>Action_Plugin_FileManage_CheckFolder_Begin 当验证文件夹是否在Z-Blog文件内前被触发</li>
          <li>Action_Plugin_FileManage_CheckFolder_End 当验证完毕后被触发</li>
          <li>Action_Plugin_FileManage_ExportSiteUpload_Begin 当加载上传页面时被触发</li>
          <li>Action_Plugin_FileManage_ExportSiteUpload_End 当加载上传页面完毕后被触发</li>
          <li>Action_Plugin_FileManage_Upload_Begin 当上传开始时被触发</li>
          <li>Action_Plugin_FileManage_Upload_End 当上传完毕后被触发</li>
          <li>Action_Plugin_FileManage_PostSiteFile_Begin 当文件保存时被触发</li>
          <li>Action_Plugin_FileManage_PostSiteFile_End 当文件保存后被触发</li>
          <li>Action_Plugin_FileManage_CreateFolder_Begin 当创建文件夹时被触发</li>
          <li>Action_Plugin_FileManage_CreateFolder_End 当创建文件夹后被触发</li>
          <li>Action_Plugin_FileManage_GetSize_Begin 当得到文件大小时被触发</li>
          <li>Action_Plugin_FileManage_GetSize_End 当得到文件大小后被触发<br />
          </li>
        </ol>
      </li>
      <li>当使用本插件时，  Action_Plugin_SiteFileEdt、  Action_Plugin_SiteFilePst 、  Action_Plugin_SiteFileDel  、  Action_Plugin_SiteFileMng的Begin和End共8个接口被替代，不再使用。</li>
      <li>Response_Plugin_SiteFileMng_SubMenu可正常使用，但插件不支持Response_Plugin_SiteFileEdt_SubMenu接口</li>
    </ol>
  </li>
  </ol>
<p>&nbsp;</p>
<p>&nbsp;</p>
  </div>
</div>
<script>

	//斑马线
	var tables=document.getElementsByTagName("ol");
	var b=false;
	for (var j = 0; j < tables.length; j++){

		var cells = tables[j].getElementsByTagName("li");

		for (var i = 0; i < cells.length; i++){
			if(b){
				cells[i].style.color="#333366";
				cells[i].style.background="#F1F4F7";
				b=false;
			}
			else{
				cells[i].style.color="#666699";
				cells[i].style.background="#FFFFFF";
				b=true;
			};
		};
	}

document.close();

</script>
<%
End Function

%>

<%
'----------------------------------------------------------
'**************  风声 ASP 无组件上传类 V2.11  *************
'作者：风声
'网站：http://www.fonshen.com
'邮件：webmaster@fonshen.com
'版权：版权全体,源代码公开,各种用途均可免费使用
'其他：有稍作改动，
'**********************************************************
'----------------------------------------------------------
Class FileManage_UpLoadClass

	Private m_TotalSize,m_MaxSize,m_FileType,m_SavePath,m_AutoSave,m_Error,m_Charset
	Private m_dicForm,m_binForm,m_binItem,m_strDate,m_lngTime
	Public	FormItem,FileItem

	Public Property Get Version
		Version="Fonshen ASP UpLoadClass Version 2.11"
	End Property

	Public Property Get Error
		Error=m_Error
	End Property

	Public Property Get Charset
		Charset=m_Charset
	End Property
	Public Property Let Charset(strCharset)
		m_Charset=strCharset
	End Property

	Public Property Get TotalSize
		TotalSize=m_TotalSize
	End Property
	Public Property Let TotalSize(lngSize)
		if isNumeric(lngSize) then m_TotalSize=Clng(lngSize)
	End Property

	Public Property Get MaxSize
		MaxSize=m_MaxSize
	End Property
	Public Property Let MaxSize(lngSize)
		if isNumeric(lngSize) then m_MaxSize=Clng(lngSize)
	End Property

	Public Property Get FileType
		FileType=m_FileType
	End Property
	Public Property Let FileType(strType)
		m_FileType=strType
	End Property

	Public Property Get SavePath
		SavePath=m_SavePath
	End Property
	Public Property Let SavePath(strPath)
		m_SavePath=Replace(strPath,chr(0),"")
	End Property

	Public Property Get AutoSave
		AutoSave=m_AutoSave
	End Property
	Public Property Let AutoSave(byVal Flag)
		select case Flag
			case 0,1,2: m_AutoSave=Flag
		end select
	End Property

	Private Sub Class_Initialize
		m_Error	   = -1
		m_Charset  = "gb2312"
		m_TotalSize= 0
		m_MaxSize  = 153600
		m_FileType = "jpg/gif"
		m_SavePath = ""
		m_AutoSave = 0
		Dim dtmNow : dtmNow = Date()
		m_strDate  = Year(dtmNow)&Right("0"&Month(dtmNow),2)&Right("0"&Day(dtmNow),2)
		m_lngTime  = Clng(Timer()*1000)
		Set m_binForm = Server.CreateObject("ADODB.Stream")
		Set m_binItem = Server.CreateObject("ADODB.Stream")
		Set m_dicForm = Server.CreateObject("Scripting.Dictionary")
		m_dicForm.CompareMode = 1
	End Sub

	Private Sub Class_Terminate
		m_dicForm.RemoveAll
		Set m_dicForm = nothing
		Set m_binItem = nothing
		m_binForm.Close()
		Set m_binForm = nothing
	End Sub

	Public Function Open()
		Open = 0
		if m_Error=-1 then
			m_Error=0
		else
			Exit Function
		end if
		Dim lngRequestSize : lngRequestSize=Request.TotalBytes
		if m_TotalSize>0 and lngRequestSize>m_TotalSize then
			m_Error=5
			Exit Function
		elseif lngRequestSize<1 then
			m_Error=4
			Exit Function
		end if

		Dim lngChunkByte : lngChunkByte = 102400
		Dim lngReadSize : lngReadSize = 0
		m_binForm.Type = 1
		m_binForm.Open()
		do
			m_binForm.Write Request.BinaryRead(lngChunkByte)
			lngReadSize=lngReadSize+lngChunkByte
			if  lngReadSize >= lngRequestSize then exit do
		loop		
		m_binForm.Position=0
		Dim binRequestData : binRequestData=m_binForm.Read()

		Dim bCrLf,strSeparator,intSeparator
		bCrLf=ChrB(13)&ChrB(10)
		intSeparator=InstrB(1,binRequestData,bCrLf)-1
		strSeparator=LeftB(binRequestData,intSeparator)

		Dim strItem,strInam,strFtyp,strPuri,strFnam,strFext,lngFsiz
		Const strSplit="'"">"
		Dim strFormItem,strFileItem,intTemp,strTemp
		Dim p_start : p_start=intSeparator+2
		Dim p_end
		Do
			p_end = InStrB(p_start,binRequestData,bCrLf&bCrLf)-1
			m_binItem.Type=1
			m_binItem.Open()
			m_binForm.Position=p_start
			m_binForm.CopyTo m_binItem,p_end-p_start
			m_binItem.Position=0
			m_binItem.Type=2
			m_binItem.Charset=m_Charset
			strItem = m_binItem.ReadText()
			m_binItem.Close()
			intTemp=Instr(39,strItem,"""")
			strInam=Mid(strItem,39,intTemp-39)

			p_start = p_end + 4
			p_end = InStrB(p_start,binRequestData,strSeparator)-1
			m_binItem.Type=1
			m_binItem.Open()
			m_binForm.Position=p_start
			lngFsiz=p_end-p_start-2
			m_binForm.CopyTo m_binItem,lngFsiz

			if Instr(intTemp,strItem,"filename=""")<>0 then
			if not m_dicForm.Exists(strInam&"_From") then
				strFileItem=strFileItem&strSplit&strInam
				if m_binItem.Size<>0 then
					intTemp=intTemp+13
					strFtyp=Mid(strItem,Instr(intTemp,strItem,"Content-Type: ")+14)
					strPuri=Mid(strItem,intTemp,Instr(intTemp,strItem,"""")-intTemp)
					intTemp=InstrRev(strPuri,"\")
					strFnam=Mid(strPuri,intTemp+1)
					m_dicForm.Add strInam&"_Type",strFtyp
					m_dicForm.Add strInam&"_Name",strFnam
					m_dicForm.Add strInam&"_Path",Left(strPuri,intTemp)
					m_dicForm.Add strInam&"_Size",lngFsiz
					if Instr(strFnam,".")<>0 then
						strFext=Mid(strFnam,InstrRev(strFnam,".")+1)
					else
						strFext=""
					end if

					select case strFtyp
					case "image/jpeg","image/pjpeg","image/jpg"
						if Lcase(strFext)<>"jpg" then strFext="jpg"
						m_binItem.Position=3
						do while not m_binItem.EOS
							do
								intTemp = Ascb(m_binItem.Read(1))
							loop while intTemp = 255 and not m_binItem.EOS
							if intTemp < 192 or intTemp > 195 then
								m_binItem.read(Bin2Val(m_binItem.Read(2))-2)
							else
								Exit do
							end if
							do
								intTemp = Ascb(m_binItem.Read(1))
							loop while intTemp < 255 and not m_binItem.EOS
						loop
						m_binItem.Read(3)
						m_dicForm.Add strInam&"_Height",Bin2Val(m_binItem.Read(2))
						m_dicForm.Add strInam&"_Width",Bin2Val(m_binItem.Read(2))
					case "image/gif"
						if Lcase(strFext)<>"gif" then strFext="gif"
						m_binItem.Position=6
						m_dicForm.Add strInam&"_Width",BinVal2(m_binItem.Read(2))
						m_dicForm.Add strInam&"_Height",BinVal2(m_binItem.Read(2))
					case "image/png"
						if Lcase(strFext)<>"png" then strFext="png"
						m_binItem.Position=18
						m_dicForm.Add strInam&"_Width",Bin2Val(m_binItem.Read(2))
						m_binItem.Read(2)
						m_dicForm.Add strInam&"_Height",Bin2Val(m_binItem.Read(2))
					case "image/bmp"
						if Lcase(strFext)<>"bmp" then strFext="bmp"
						m_binItem.Position=18
						m_dicForm.Add strInam&"_Width",BinVal2(m_binItem.Read(4))
						m_dicForm.Add strInam&"_Height",BinVal2(m_binItem.Read(4))
					case "application/x-shockwave-flash"
						if Lcase(strFext)<>"swf" then strFext="swf"
						m_binItem.Position=0
						if Ascb(m_binItem.Read(1))=70 then
							m_binItem.Position=8
							strTemp = Num2Str(Ascb(m_binItem.Read(1)), 2 ,8)
							intTemp = Str2Num(Left(strTemp, 5), 2)
							strTemp = Mid(strTemp, 6)
							while (Len(strTemp) < intTemp * 4)
								strTemp = strTemp & Num2Str(Ascb(m_binItem.Read(1)), 2 ,8)
							wend
							m_dicForm.Add strInam&"_Width", Int(Abs(Str2Num(Mid(strTemp, intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 1, intTemp), 2)) / 20)
							m_dicForm.Add strInam&"_Height",Int(Abs(Str2Num(Mid(strTemp, 3 * intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 2 * intTemp + 1, intTemp), 2)) / 20)
						end if
					end select

					m_dicForm.Add strInam&"_Ext",strFext
					m_dicForm.Add strInam&"_From",p_start
					if m_AutoSave<>2 then
						intTemp=GetFerr(lngFsiz,strFext)
						m_dicForm.Add strInam&"_Err",intTemp
						if intTemp=0 then
							if m_AutoSave=0 then
								strFnam=GetTimeStr()
								if strFext<>"" then strFnam=strFnam&"."&strFext
							end if
							m_binItem.SaveToFile m_SavePath&strFnam,2
							m_dicForm.Add strInam,strFnam
						end if
					end if
				else
					m_dicForm.Add strInam&"_Err",-1
				end if
			end if
			else
				m_binItem.Position=0
				m_binItem.Type=2
				m_binItem.Charset=m_Charset
				strTemp=m_binItem.ReadText
				if m_dicForm.Exists(strInam) then
					m_dicForm(strInam) = m_dicForm(strInam)&","&strTemp
				else
					strFormItem=strFormItem&strSplit&strInam
					m_dicForm.Add strInam,strTemp
				end if
			end if

			m_binItem.Close()
			p_start = p_end+intSeparator+2
		loop Until p_start+3>lngRequestSize
		FormItem=Split(strFormItem,strSplit)
		FileItem=Split(strFileItem,strSplit)
		
		Open = lngRequestSize
	End Function

	Private Function GetTimeStr()
		m_lngTime=m_lngTime+1
		GetTimeStr=m_strDate&Right("00000000"&m_lngTime,8)
	End Function

	Private Function GetFerr(lngFsiz,strFext)
		dim intFerr
		intFerr=0
		if lngFsiz>m_MaxSize and m_MaxSize>0 then
			if m_Error=0 or m_Error=2 then m_Error=m_Error+1
			intFerr=intFerr+1
		end if
		if Instr(1,LCase("/"&m_FileType&"/"),LCase("/"&strFext&"/"))=0 and m_FileType<>"" then
			if m_Error<2 then m_Error=m_Error+2
			intFerr=intFerr+2
		end if
		GetFerr=intFerr
	End Function

	Public Function Save(Item,strFnam)
		Save=false
		if m_dicForm.Exists(Item&"_From") then
			dim intFerr,strFext
			strFext=m_dicForm(Item&"_Ext")
			intFerr=GetFerr(m_dicForm(Item&"_Size"),strFext)
			if m_dicForm.Exists(Item&"_Err") then
				if intFerr=0 then
					m_dicForm(Item&"_Err")=0
				end if
			else
				m_dicForm.Add Item&"_Err",intFerr
			end if
			if intFerr<>0 then Exit Function
			if VarType(strFnam)=2 then
				select case strFnam
					case 0:strFnam=GetTimeStr()
						if strFext<>"" then strFnam=strFnam&"."&strFext
					case 1:strFnam=m_dicForm(Item&"_Name")
				end select
			end if
			m_binItem.Type = 1
			m_binItem.Open
			m_binForm.Position = m_dicForm(Item&"_From")
			m_binForm.CopyTo m_binItem,m_dicForm(Item&"_Size")
			m_binItem.SaveToFile m_SavePath&strFnam,2
			m_binItem.Close()
			if m_dicForm.Exists(Item) then
				m_dicForm(Item)=strFnam
			else
				m_dicForm.Add Item,strFnam
			end if
			Save=true
		end if
	End Function

	Public Function GetData(Item)
		GetData=""
		if m_dicForm.Exists(Item&"_From") then
			if GetFerr(m_dicForm(Item&"_Size"),m_dicForm(Item&"_Ext"))<>0 then Exit Function
			m_binForm.Position = m_dicForm(Item&"_From")
			GetData = m_binForm.Read(m_dicForm(Item&"_Size"))
		end if
	End Function

	Public Function Form(Item)
		if m_dicForm.Exists(Item) then
			Form=m_dicForm(Item)
		else
			Form=""
		end if
	End Function

	Private Function BinVal2(bin)
		dim lngValue,i
		lngValue=0
		for i = lenb(bin) to 1 step -1
			lngValue = lngValue *256 + Ascb(midb(bin,i,1))
	Next
		BinVal2=lngValue
	End Function

	Private Function Bin2Val(bin)
		dim lngValue,i
		lngValue=0
		for i = 1 to lenb(bin)
			lngValue = lngValue *256 + Ascb(midb(bin,i,1))
	Next
		Bin2Val=lngValue
	End Function

	Private Function Num2Str(num, base, lens)
		Dim ret,i
		ret = ""
		while(num >= base)
			i   = num Mod base
			ret = i & ret
			num = (num - i) / base
		wend
		Num2Str = Right(String(lens, "0") & num & ret, lens)
	End Function

	Private Function Str2Num(str, base)
		Dim ret, i
		ret = 0 
		for i = 1 to Len(str)
			ret = ret * base + Cint(Mid(str, i, 1))
	Next
		Str2Num = ret
	End Function

End Class
%>