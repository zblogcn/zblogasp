<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Pre Terminator 及以上版本, 其它版本的Z-blog未知
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    主题管理插件
'// 最后修改：   2008-7-6
'// 最后版本:    1.2
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Server.ScriptTimeout=99999999 %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../function/c_function.asp" -->
<!-- #include file="../../function/c_system_lib.asp" -->
<!-- #include file="../../function/c_system_base.asp" -->
<!-- #include file="../../function/c_system_plugin.asp" -->
<!-- #include file="c_sapper.asp" -->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("ThemeSapper")=False Then Call ShowError(48)

BlogTitle = "将主题打包"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta name="robots" content="noindex,nofollow"/>
	<link rel="stylesheet" rev="stylesheet" href="../../CSS/admin.css" type="text/css" media="screen" />
	<link rel="stylesheet" rev="stylesheet" href="images/style.css" type="text/css" media="screen" />
	<title><%=BlogTitle%></title>
</head>
<body>
<div id="divMain">
	<div class="Header">Theme Sapper - 主题导出 - 导出主题为 ZTI 文件. <a href="help.asp#exportzti" title="如何导出主题为 ZTI 文件">[页面帮助]</a></div>
	<%Call SapperMenu("0")%>
<div id="divMain2">
<%
Action=Request.QueryString("act")
SelectedTheme=Request.QueryString("theme")
SelectedThemeName=Request.QueryString("themename")

If Action <> "pack" Then Call GetBlogHint()
Response.Write "<div>"
Response.Flush

If Action="view" Then
	Call ViewXMLPackInfo()
End If


If Action="" Then
	Response.Write "<p id=""loading"">正在载入主题信息, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Call EditXMLPackInfo()

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('loading').style.display = 'none';</script>"
End If


If Action="pack" Then

	Dim Pack_Error
	Pack_Error=0

	If SelectedTheme="" Then
		Response.Write "<p><font color=""red""> × 主题的名称为空.</font></p>"
		Pack_Error=Pack_Error+1

	Else
		Response.Write "<p id=""loading2"">正在打包主题, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
		Response.Write "<p class=""status-box"">正在打包主题...<p>"
		Response.Flush

		Dim ZipPathDir,ZipPathFile,Pack_ThemeDir
		Dim TS_startime,TS_endtime

		'打包文件目录与生成文件名
		ZipPathDir = BlogPath & "THEMES\" & LCase(SelectedTheme)
		If right(ZipPathDir,1)<>"\" Then ZipPathDir=ZipPathDir&"\"

		Pack_ThemeDir = SelectedTheme
		If right(Pack_ThemeDir,1)<>"\" Then Pack_ThemeDir=Pack_ThemeDir &"\"


		If Request.Form("PubOrBack")="Pub" Then 
			ZipPathFile = BlogPath & "PLUGIN\ThemeSapper\Export\" & SelectedTheme & ".zti"
			Pack_Error=Pack_Error+DeleteFile(ZipPathFile)
		ElseIf Request.Form("PubOrBack")="Bak" Then
			Dim BakNum
			BakNum = 0
			Do
				BakNum = BakNum + 1
				ZipPathFile=BlogPath & "PLUGIN\ThemeSapper\Export\" & SelectedTheme & "_Bak("& Cstr(BakNum) & ").zti"
			Loop Until FileExists(ZipPathFile)=False
		Else
			Response.Write "<p><font color=""red""> × 导出设置错误.</font></p>"
			ZipPathFile = BlogPath & "PLUGIN\ThemeSapper\Export\" & SelectedTheme & ".xml"
			Pack_Error=Pack_Error+1
		End If

		'开始打包
		CreateXml(ZipPathFile)
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('loading2').style.display = 'none';</script>"
	End If

	If Pack_Error = 0 Then
		If Request.Form("PubOrBack")="Pub" Then 
			Response.Write "<p class=""status-box""> √ 主题打包完成, 并保存在主机上, 名称为: """ & SelectedTheme & ".zti"". 如果您的浏览器没能自动跳转, 请 <a href=""Xml_Pack.asp?act=view&amp;theme="& Server.URLEncode(SelectedTheme) &"&amp;themename="& Server.URLEncode(SelectedTheme) &""">[点击这里]</a>.</p>"
			Response.Write "<script>setTimeout(""self.location.href='Xml_Pack.asp?act=view&theme="& Server.URLEncode(SelectedTheme) &"&themename="& Server.URLEncode(SelectedTheme) &"'"",3000);</script>"
		ElseIf Request.Form("PubOrBack")="Bak" Then
			Response.Write "<p class=""status-box""> √ 主题打包完成, 并保存在主机上, 名称为: """ & SelectedTheme & "_Bak("& Cstr(BakNum) & ").zti"". 如果您的浏览器没能自动跳转, 请 <a href=""Xml_Pack.asp?act=view&amp;theme="& Server.URLEncode(SelectedTheme & "_Bak("& Cstr(BakNum) & ")") &"&amp;themename="& Server.URLEncode(SelectedTheme) &""">[点击这里]</a>.</p>"
			Response.Write "<script>setTimeout(""self.location.href='Xml_Pack.asp?act=view&theme="& Server.URLEncode(SelectedTheme & "_Bak("& Cstr(BakNum) & ")") &"&themename="& Server.URLEncode(SelectedTheme) &"'"",3000);</script>"
		Else
			Response.Write "<p class=""status-box""><font color=""red""> × 这种情况不可能发生. </font></p>"
		End If
	Else
		Response.Write "<p class=""status-box""><font color=""red""> × 主题打包失败. "
		Response.Write "<a href=""javascript:history.back(-1)"" title=""返回上一个页面""><span>[返回]</span></a> 或 <a href=""javascript:window.location.reload();"" title=""点此重试""><span>[重试]</span></a></font></p>"
	End If

End If

Response.Write "</div>"
%>
</div>
</div>
</body>
</html>
<%

'预览XML安装包内的信息
Sub ViewXMLPackInfo()
On Error Resume Next

	If InStr(UCase(Request.ServerVariables("HTTP_USER_AGENT")),UCase("Opera"))>0 Then '如果是Opera浏览器
		Response.Write "<p class=""download-box""><a href=""Export/"& SelectedTheme & ".zti"" title=""右键另存为下载此 ZTI 文件"">[右键点击这里, 选择 ""链接另存为...(Save Link As...)""  保存此 ZTI 文件 - " & SelectedTheme & ".zti - 到本地]</a></p>"
	Else
		Response.Write "<p class=""download-box""><a href=""Xml_Download.asp?theme="& Server.URLEncode(SelectedTheme & ".zti") &""" title=""左键点击下载此 ZTI 文件"">[左键点击这里下载此 ZTI 文件 - " & SelectedTheme & ".zti - 到本地]</a>"
	End If

	Dim objXmlFile,strXmlFile
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

		If fso.FileExists(BlogPath & "PLUGIN/ThemeSapper/Export/" & SelectedTheme & ".zti") Then

			strXmlFile =BlogPath & "PLUGIN/ThemeSapper/Export/" & SelectedTheme & ".zti"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else

					ThemeAuthor_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					ThemeAuthor_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
					ThemeAuthor_Email=objXmlFile.documentElement.selectSingleNode("author/email").text

					ThemeID=objXmlFile.documentElement.selectSingleNode("id").text
					ThemeName=objXmlFile.documentElement.selectSingleNode("name").text
					ThemeURL=objXmlFile.documentElement.selectSingleNode("url").text
					ThemePubDate=objXmlFile.documentElement.selectSingleNode("pubdate").text
					ThemeNote=objXmlFile.documentElement.selectSingleNode("note").text

					ThemeAdapted=objXmlFile.documentElement.selectSingleNode("adapted").text
					ThemeVersion=objXmlFile.documentElement.selectSingleNode("version").text
					ThemeModified=objXmlFile.documentElement.selectSingleNode("modified").text

				End If
			Set objXmlFile=Nothing
			End If

			If fso.FileExists(BlogPath & "/THEMES/" & SelectedThemeName & "/" & "screenshot.png") Then
				ThemeScreenShot="../../THEMES/" & SelectedThemeName & "/" & "screenshot.png"
			Else
				ThemeScreenShot="Images/noscreenshot.png"
			End If

			Response.Write "<div class=""themeDetail"">"

			Response.Write "<p><img src=""" & ThemeScreenShot & """ title=""" & ThemeName & """ alt=""ScreenShot"" /></p>"

			Response.Write "<p>以下为 ZTI 文件中所包含的信息:</p><hr />"

			Response.Write "<p><b>ID:</b> " & ThemeID & "</p>"
			Response.Write "<p><b>名称:</b> " & ThemeName & "</p>"
			If ThemeURL<>Empty Then Response.Write "<p><b>网址:</b> <a href=""" & ThemeURL & """ target=""_blank"" title=""主题发布地址"">" & ThemeURL & "</a></p>"
			If ThemeAuthor_Url=Empty Then
				Response.Write "<p><b>作者:</b> " & ThemeAuthor_Name & "</p>"
			Else
				Response.Write "<p><b>作者:</b> <a href=""" & ThemeAuthor_Url & """ target=""_blank"" title=""作者主页"">" & ThemeAuthor_Name & "</a></p>"
			End If
			If ThemeAuthor_Email<>Empty Then Response.Write "<p><b>邮箱:</b> <a href=""mailto:" & ThemeAuthor_Email & """ title=""作者邮箱"">" & ThemeAuthor_Email & "</a></p>"
			Response.Write "<p><b>发布:</b> " & ThemePubDate & "</p>"
			Response.Write "<p><b>简介:</b> " & ThemeNote & "</p><br />"

			Response.Write "<p><b>适用:</b> " & ThemeAdapted & "</p>"
			Response.Write "<p><b>版本:</b> " & ThemeVersion & "</p>"
			Response.Write "<p><b>修正:</b> " & ThemeModified & "</p><hr />"

			Response.Write "</div>"

			Response.Write "<p><form id=""edit"" name=""edit"" method=""get"" action=""#"">"
			Response.Write "<p><input onclick=""self.location.href='ThemeList.asp';"" type=""button"" class=""button"" value=""返回主题管理"" title=""返回主题管理页"" /></p>"
			Response.Write "</form></p>"

		Else
			Response.Write "<p><font color=""red""> × 无法找到主题包. </p>"
		End If
	Set fso = nothing
	Err.Clear

End Sub


'编辑XML安装包内的信息
Sub EditXMLPackInfo()
On Error Resume Next

	Dim objXmlFile,strXmlFile
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

		If fso.FileExists(BlogPath & "/THEMES/" & SelectedTheme & "/" & "theme.xml") Then

			strXmlFile =BlogPath & "/THEMES/" & SelectedTheme & "/" & "theme.xml"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else

					ThemeAuthor_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					ThemeAuthor_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
					ThemeAuthor_Email=objXmlFile.documentElement.selectSingleNode("author/email").text

					ThemeSource_Name=objXmlFile.documentElement.selectSingleNode("source/name").text
					ThemeSource_Url=objXmlFile.documentElement.selectSingleNode("source/url").text
					ThemeSource_Email=objXmlFile.documentElement.selectSingleNode("source/email").text

					If ThemeAuthor_Name=Empty Then
						ThemeAuthor_Name=ThemeSource_Name
						ThemeAuthor_Url=ThemeSource_Url
						ThemeAuthor_Email=ThemeSource_Email
					End If

					ThemeID=objXmlFile.documentElement.selectSingleNode("id").text
					ThemeName=objXmlFile.documentElement.selectSingleNode("name").text
					ThemeURL=objXmlFile.documentElement.selectSingleNode("url").text
					ThemePubDate=objXmlFile.documentElement.selectSingleNode("pubdate").text
					ThemeNote=objXmlFile.documentElement.selectSingleNode("note").text

					ThemeAdapted=objXmlFile.documentElement.selectSingleNode("adapted").text
					ThemeVersion=objXmlFile.documentElement.selectSingleNode("version").text
					ThemeModified=objXmlFile.documentElement.selectSingleNode("modified").text

					ThemeDescription=objXmlFile.documentElement.selectSingleNode("description").text

					ThemeAuthor_Name=TransferHTML(ThemeAuthor_Name,"[html-format]")
					ThemeSource_Name=TransferHTML(ThemeSource_Name,"[html-format]")
					ThemeName=TransferHTML(ThemeName,"[html-format]")
					ThemeNote=TransferHTML(ThemeNote,"[html-format]")
					ThemeDescription=TransferHTML(ThemeDescription,"[html-format]")


				End If
			Set objXmlFile=Nothing
			End If

		Else

			ThemeSource_Name="unknown"
			ThemeSource_Url=Empty
			ThemeSource_Email="null@null.com"

			ThemeAuthor_Name="unknown"
			ThemeAuthor_Url=Empty
			ThemeAuthor_Email="null@null.com"

			ThemeName=SelectedTheme
			ThemeURL=Empty
			ThemeNote="unknown"
			ThemePubDate=Date()

			ThemeAdapted="Z-Blog 1.8"
			ThemeVersion="1.0"
			ThemeModified=Date()

			ThemeDescription="nothing..."

		End If
	Set fso = nothing
	Err.Clear

	Response.Write "<form id=""edit"" name=""edit"" method=""post"">"


	Response.Write "<p><b>请指定 ZTI 文件中所包含的主题信息, 仅用于基于 Z-blog 1.8 的主题. <a href=""help.asp#aboutzti"">[什么是 ZTI 文件?]</a></b></p><hr />"

	Response.Write "<p>※主题ID: (插件ID应为插件信息文档中的ID, 此处不可修改.)</p><p><input name=""ThemeID"" style=""width:99%"" type=""text"" value="""&ThemeID&""" readonly /></p><p></p>"

	Response.Write "<p>※主题名称:</p><p><input name=""ThemeName"" style=""width:99%"" type=""text"" value="""&ThemeName&""" /></p><p></p>"

	Response.Write "<p>※主题的发布页面: (强列建议填写, 以方便使用者在安装插件前能看到作者的发布页面, 从而获得更多的发布信息.)</p><p><input name=""ThemeURL"" style=""width:99%"" type=""text"" value="""&ThemeURL&""" /></p><p></p>"

	Response.Write "<p>※主题简介 (可以用 &lt;br /&gt; 换行, 可以使用 html 标签):</p><p><textarea name=""ThemeNote"" style=""width:99%"" rows=""5"">"&ThemeNote&"</textarea></p><p></p>"

	Response.Write "<p><b>以下信息从主题信息文档 Theme.xml 中读取且必须与之保持一致, 此处不可修改. <a href=""Xml_Edit.asp?theme=" & Server.URLEncode(SelectedTheme) & """ title=""编辑主题信息文档-Theme.xml"">[编辑主题信息文档]</a></b></p><br />"


	Response.Write "<p>※主题适用的 Z-Blog 版本: (写法应为: ""Z-Blog 1.8 Spirit"")</p><p><input name=""ThemeAdapted"" style=""width:99%"" type=""text"" value="""&ThemeAdapted&""" readonly /></p><p></p>"

	Response.Write "<p>※主题的修订版本号:</p><p><input name=""ThemeVersion"" style=""width:99%"" type=""text"" value="""&ThemeVersion&""" readonly /></p><p></p>"

	Response.Write "<p>※主题的发布日期: (日期标准格式应为:"""&Date()&""")</p><p><input name=""ThemePubDate"" style=""width:99%"" type=""text"" value="""&ThemePubDate&""" readonly /></p><p></p>"

	Response.Write "<p>※主题的最后修改日期: (日期标准格式应为:"""&Date()&""")</p><p><input name=""ThemeModified"" style=""width:99%"" type=""text"" value="""&ThemeModified&""" readonly /></p><p></p>"

	Response.Write "<p>※主题作者:</p><p><input name=""AuthorName"" style=""width:99%"" type=""text"" value="""&ThemeAuthor_Name&""" readonly /></p><p></p>"

	Response.Write "<p>※主题作者主页:</p><p><input name=""AuthorURL"" style=""width:99%"" type=""text"" value="""&ThemeAuthor_Url&""" readonly /></p><p></p>"

	Response.Write "<p>※主题作者 Email:</p><p><input name=""AuthorEmail"" style=""width:99%"" type=""text"" value="""&ThemeAuthor_Email&""" readonly /></p><p></p>"

	Response.Write "<p><b>主题将被按 Z-Blog 主题专用安装包 Version 0.1 打包成 ZTI 文件, 并保存在 TS 插件的 Export 目录下.</b></p><hr />"

	Response.Write "<p><input name=""PubOrBack"" type=""radio"" value=""Pub"" checked=""checked""/> 这次导出是为了<b>发布</b> (导出的文件名必须为 <b>"& SelectedTheme &".zti</b>)<br /><input name=""PubOrBack"" type=""radio"" value=""Bak""/> 这次导出是为了<b>备份</b> (导出的文件名为 <b>"& SelectedTheme &"_Bak(n).zti</b> 的形式)</p><hr />"


	Response.Write "<p><input type=""submit"" class=""button"" value=""确认信息并打包主题"" id=""btnPost"" onclick='document.getElementById(""edit"").action=""Xml_Pack.asp?act=pack&theme="& SelectedTheme &""";' title=""确认信息并打包主题"" /> <input onclick=""self.location.href='ThemeList.asp';"" type=""button"" class=""button"" value=""取消并返回主题管理"" title=""取消并返回主题管理页"" />  <input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /></p>"


	Response.Write "</form>"

End Sub

'遍历目录内的所有文件以及文件夹
Sub LoadData(DirPath)
On Error Resume Next

	Dim XmlDoc
	Dim fso            'fso对象
	Dim objFolder      '文件夹对象
	Dim objSubFolders  '子文件夹集合
	Dim objSubFolder   '子文件夹对象
	Dim objFiles       '文件集合
	Dim objFile        '文件对象
	Dim objStream
	Dim pathname,TextStream,pp,Xfolder,Xfpath,Xfile,Xpath,Xstream
	Dim PathNameStr

	Set fso=server.CreateObject("scripting.filesystemobject")
	Set objFolder=fso.GetFolder(DirPath)'创建文件夹对象
	
	Response.Write("<p>打包目录: "& Replace(DirPath,BlogPath,"") &"...</p>")
	Response.flush
	
	Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
	XmlDoc.async = False
	XmlDoc.ValidateOnParse=False
	XmlDoc.load (ZipPathFile)

	'写入每个文件夹路径
	Set Xfolder = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("folder"))
	Set Xfpath = Xfolder.AppendChild(XmlDoc.CreateElement("path"))
		Xfpath.text = replace(DirPath,ZipPathDir,Pack_ThemeDir)

	Response.Write "<blockquote><font color=""Teal"">"
		Set objFiles=objFolder.Files
			for each objFile in objFiles
				If lcase(DirPath & objFile.name) <> lcase(Request.ServerVariables("PATH_TRANSLATED")) Then
					PathNameStr = DirPath & "" & objFile.name
					Response.Write Replace(PathNameStr,BlogPath,"") & "<br />"
					Response.flush
					'================================================
					'写入文件的路径及文件内容
				   Set Xfile = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("file"))
				   Set Xpath = Xfile.AppendChild(XmlDoc.CreateElement("path"))
					   Xpath.text = replace(PathNameStr,ZipPathDir,Pack_ThemeDir)
				   '创建文件流读入文件内容，并写入XML文件中
				   Set objStream = Server.CreateObject("ADODB.Stream")
				   objStream.Type = 1
				   objStream.Open()
				   objStream.LoadFromFile(PathNameStr)
				   objStream.position = 0
				   
				   Set Xstream = Xfile.AppendChild(XmlDoc.CreateElement("stream"))
					   Xstream.SetAttribute "xmlns:dt","urn:schemas-microsoft-com:datatypes"
					   '文件内容采用二制方式存放
					   Xstream.dataType = "bin.base64"
					   Xstream.nodeTypedValue = objStream.Read()
				   
				   Set objStream=Nothing
				   Set Xpath = Nothing
				   Set Xstream = Nothing
				   Set Xfile = Nothing
				  '================================================
				end if
			next
	Response.Write "</font></blockquote>"
	XmlDoc.Save(ZipPathFile)
	Set Xfpath = Nothing
	Set Xfolder = Nothing
	Set XmlDoc = Nothing
	
	'创建的子文件夹对象
	Set objSubFolders=objFolder.Subfolders
		'调用递归遍历子文件夹
		for each objSubFolder in objSubFolders
			pathname = DirPath & objSubFolder.name & "\"
			LoadData(pathname)
		next
	Set objFolder=Nothing
	Set objSubFolders=Nothing
	Set fso=Nothing

	If Err.Number<>0 Then Pack_Error=Pack_Error+1
	Err.Clear

End Sub



'创建一个空的XML文件，为写入文件作准备
Sub CreateXml(FilePath)
On Error Resume Next

	'程序开始执行时间
	TS_startime=timer()
	Dim XmlDoc,Root,xRoot
	Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
		XmlDoc.async = False
		XmlDoc.ValidateOnParse=False
		Set Root = XmlDoc.createProcessingInstruction("xml","version='1.0' encoding='utf-8'")
		XmlDoc.appendChild(Root)
		Set xRoot = XmlDoc.appendChild(XmlDoc.CreateElement("root"))
			xRoot.setAttribute "version",XML_Pack_Ver
			xRoot.setAttribute "type",XML_Pack_Type
			xRoot.setAttribute "for",XML_Pack_Version
		Set xRoot = Nothing

		'写入文件信息
		Dim Author,AuthorName,AuthorURL,AuthorEmail

		Set ThemeID = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("id"))
			ThemeID.Text = SelectedTheme
		Set ThemeID=Nothing

		Set ThemeName = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("name"))
			ThemeName.Text = Request.Form("ThemeName")
		Set ThemeName=Nothing

		Set ThemeURL = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("url"))
			ThemeURL.Text = Request.Form("ThemeURL")
		Set ThemeURL=Nothing

		Set ThemePubDate = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("pubdate"))
			ThemePubDate.Text = Request.Form("ThemePubDate")
		Set ThemePubDate=Nothing

		Set ThemeAdapted = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("adapted"))
			ThemeAdapted.Text = Request.Form("ThemeAdapted")
		Set ThemeAdapted=Nothing

		Set ThemeVersion = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("version"))
			ThemeVersion.Text = Request.Form("ThemeVersion")
		Set ThemeVersion=Nothing

		Set ThemeModified = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("modified"))
			ThemeModified.Text = Request.Form("ThemeModified")
		Set ThemeModified=Nothing

		Set ThemeNote = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("note"))
			ThemeNote.Text = Replace(Replace(Request.Form("ThemeNote"),vbCr,""),vbLf,"")
		Set ThemeNote=Nothing

		Set Author = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("author"))

			Set AuthorName = Author.AppendChild(XmlDoc.CreateElement("name"))
				AuthorName.Text = Request.Form("AuthorName")
			Set AuthorName=Nothing

			Set AuthorURL = Author.AppendChild(XmlDoc.CreateElement("url"))
				AuthorURL.Text = Request.Form("AuthorURL")
			Set AuthorURL=Nothing

			Set AuthorEmail = Author.AppendChild(XmlDoc.CreateElement("email"))
				AuthorEmail.Text = Request.Form("AuthorEmail")
			Set AuthorEmail=Nothing

		Set Author=Nothing

		XmlDoc.Save(FilePath)
		Set Root = Nothing
	Set XmlDoc = Nothing

	If Err.Number<>0 Then Pack_Error=Pack_Error+1
	Err.Clear

	LoadData(ZipPathDir)
	'程序结束时间
	TS_endtime=timer()

	Dim TS_PageTime
	TS_PageTime=FormatNumber((TS_endtime-TS_startime),3)
	If left(TS_PageTime,1)="." Then TS_PageTime="0" & TS_PageTime

	Response.Write("<p>页面执行时间：" & TS_PageTime & "秒</p>")
End Sub

Call System_Terminate()

If Err.Number<>0 Then
	Call ShowError(0)
End If
%>