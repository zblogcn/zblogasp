<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Pre Terminator 及以上版本, 其它版本的Z-blog未知
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    插件管理插件
'// 最后修改：   2008-6-28
'// 最后版本:    1.2
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Server.ScriptTimeout=99999999 %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="c_sapper.asp" -->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("PluginSapper")=False Then Call ShowError(48)

BlogTitle = "将插件打包"
PS_Head
%>


<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
	<div class="divHeader">Plugin Sapper - 插件导出 - 导出插件为 ZPI 文件. <a href="help.asp#exportzpi" title="如何导出插件为 ZPI 文件">[页面帮助]</a></div>
	<%Call SapperMenu("0")%>
<div id="divMain2">
<%
Action=Request.QueryString("act")
SelectedPlugin=Request.QueryString("plugin")
SelectedPluginName=Request.QueryString("pluginname")

If Action <> "pack" Then Call GetBlogHint()
Response.Write "<div>"
Response.Flush

If Action="view" Then
	Call ViewXMLPackInfo()
End If


If Action="" Then
	Response.Write "<p id=""loading"">正在载入插件信息, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Call EditXMLPackInfo()

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('loading').style.display = 'none';</script>"
End If


If Action="pack" Then

	Dim Pack_Error
	Pack_Error=0

	If SelectedPlugin="" Then
		Response.Write "<p><font color=""red""> × 插件的名称为空.</font></p>"
		Pack_Error=Pack_Error+1

	Else
		Response.Write "<p id=""loading2"">正在打包插件, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
		Response.Write "<p class=""status-box"">正在打包插件...<p>"
		Response.Flush

		Dim ZipPathDir,ZipPathFile,Pack_PluginDir
		Dim TS_startime,TS_endtime

		'打包文件目录与生成文件名
		ZipPathDir = BlogPath & "ZB_USERS\PLUGIN\" & LCase(SelectedPlugin)
		If right(ZipPathDir,1)<>"\" Then ZipPathDir=ZipPathDir&"\"

		Pack_PluginDir = SelectedPlugin
		If right(Pack_PluginDir,1)<>"\" Then Pack_PluginDir=Pack_PluginDir &"\"


		If Request.Form("PubOrBack")="Pub" Then 
			ZipPathFile = BlogPath & "ZB_USERS\PLUGIN\PluginSapper\Export\" & SelectedPlugin & ".zpi"
			Pack_Error=Pack_Error+DeleteFile(ZipPathFile)
		ElseIf Request.Form("PubOrBack")="Bak" Then
			Dim BakNum
			BakNum = 0
			Do
				BakNum = BakNum + 1
				ZipPathFile=BlogPath & "ZB_USERS\PLUGIN\PluginSapper\Export\" & SelectedPlugin & "_Bak("& Cstr(BakNum) & ").zpi"
			Loop Until FileExists(ZipPathFile)=False
		Else
			Response.Write "<p><font color=""red""> × 导出设置错误.</font></p>"
			ZipPathFile = BlogPath & "ZB_USERS\PLUGIN\PluginSapper\Export\" & SelectedPlugin & ".xml"
			Pack_Error=Pack_Error+1
		End If

		'开始打包
		CreateXml(ZipPathFile)
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('loading2').style.display = 'none';</script>"
	End If

	If Pack_Error = 0 Then
		If Request.Form("PubOrBack")="Pub" Then 
			Response.Write "<p class=""status-box""> √ 插件打包完成, 并保存在主机上, 名称为: """ & SelectedPlugin & ".zpi"". 如果您的浏览器没能自动跳转, 请 <a href=""Xml_Pack.asp?act=view&amp;plugin="& Server.URLEncode(SelectedPlugin) &"&amp;pluginname="& Server.URLEncode(SelectedPlugin) &""">[点击这里]</a>.</p>"
			Response.Write "<script>setTimeout(""self.location.href='Xml_Pack.asp?act=view&plugin="& Server.URLEncode(SelectedPlugin) &"&pluginname="& Server.URLEncode(SelectedPlugin) &"'"",3000);</script>"
		ElseIf Request.Form("PubOrBack")="Bak" Then
			Response.Write "<p class=""status-box""> √ 插件打包完成, 并保存在主机上, 名称为: """ & SelectedPlugin & "_Bak("& Cstr(BakNum) & ").zpi"". 如果您的浏览器没能自动跳转, 请 <a href=""Xml_Pack.asp?act=view&amp;plugin="& Server.URLEncode(SelectedPlugin & "_Bak("& Cstr(BakNum) & ")") &"&amp;pluginname="& Server.URLEncode(SelectedPlugin) &""">[点击这里]</a>.</p>"
			Response.Write "<script>setTimeout(""self.location.href='Xml_Pack.asp?act=view&plugin="& Server.URLEncode(SelectedPlugin & "_Bak("& Cstr(BakNum) & ")") &"&pluginname="& Server.URLEncode(SelectedPlugin) &"'"",3000);</script>"
		Else
			Response.Write "<p class=""status-box""><font color=""red""> × 这种情况不可能发生. </font></p>"
		End If
	Else
		Response.Write "<p class=""status-box""><font color=""red""> × 插件打包失败. "
		Response.Write "<a href=""javascript:history.back(-1)"" title=""返回上一个页面""><span>[返回]</span></a> 或 <a href=""javascript:window.location.reload();"" title=""重试一下""><span>[重试]</span></a></font></p>"
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
		Response.Write "<p class=""download-box""><a href=""Export/"& SelectedPlugin & ".ZPI"" title=""右键另存为下载此 ZPI 文件"">[右键点击这里, 选择 ""链接另存为...(Save Link As...)""  保存此 ZPI 文件 - " & SelectedPlugin & ".zpi - 到本地]</a></p>"
	Else
		Response.Write "<p class=""download-box""><a href=""Xml_Download.asp?plugin="& Server.URLEncode(SelectedPlugin & ".zpi") &""" title=""左键点击下载此 ZPI 文件"">[左键点击这里下载此 ZPI 文件 - " & SelectedPlugin & ".zpi - 到本地]</a>"
	End If

	Dim objXmlFile,strXmlFile
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

		If fso.FileExists(BlogPath & "ZB_USERS/PLUGIN/PluginSapper/Export/" & SelectedPlugin & ".zpi") Then

			strXmlFile =BlogPath & "ZB_USERS/PLUGIN/PluginSapper/Export/" & SelectedPlugin & ".zpi"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else

					Plugin_Author_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					Plugin_Author_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
					Plugin_Author_Email=objXmlFile.documentElement.selectSingleNode("author/email").text

					Plugin_ID=objXmlFile.documentElement.selectSingleNode("id").text
					Plugin_Name=objXmlFile.documentElement.selectSingleNode("name").text
					Plugin_URL=objXmlFile.documentElement.selectSingleNode("url").text
					Plugin_PubDate=objXmlFile.documentElement.selectSingleNode("pubdate").text
					Plugin_Note=objXmlFile.documentElement.selectSingleNode("note").text

					Plugin_Adapted=objXmlFile.documentElement.selectSingleNode("adapted").text
					Plugin_Version=objXmlFile.documentElement.selectSingleNode("version").text
					Plugin_Modified=objXmlFile.documentElement.selectSingleNode("modified").text

				End If
			Set objXmlFile=Nothing
			End If

			Response.Write "<div class=""PluginDetail"">"

			Response.Write "<p>以下为 ZPI 文件中所包含的信息:</p><hr />"

			Response.Write "<p><b>ID:</b> " & Plugin_ID & "</p>"
			Response.Write "<p><b>名称:</b> " & Plugin_Name & "</p>"
			If Plugin_URL<>Empty Then Response.Write "<p><b>网址:</b> <a href=""" & Plugin_URL & """ target=""_blank"" title=""插件发布地址"">" & Plugin_URL & "</a></p>"
			If Plugin_Author_Url=Empty Then
				Response.Write "<p><b>作者:</b> " & Plugin_Author_Name & "</p>"
			Else
				Response.Write "<p><b>作者:</b> <a href=""" & Plugin_Author_Url & """ target=""_blank"" title=""作者主页"">" & Plugin_Author_Name & "</a></p>"
			End If
			If Plugin_Author_Email<>Empty Then Response.Write "<p><b>邮箱:</b> <a href=""mailto:" & PluginAuthor_Email & """ title=""作者邮箱"">" & Plugin_Author_Email & "</a></p>"
			Response.Write "<p><b>发布:</b> " & Plugin_PubDate & "</p>"
			Response.Write "<p><b>简介:</b> " & Plugin_Note & "</p><br />"

			Response.Write "<p><b>适用:</b> " & Plugin_Adapted & "</p>"
			Response.Write "<p><b>版本:</b> " & Plugin_Version & "</p>"
			Response.Write "<p><b>修正:</b> " & Plugin_Modified & "</p><hr />"

			Response.Write "</div>"

			Response.Write "<p><form id=""edit"" name=""edit"" method=""get"" action=""#"">"
			Response.Write "<p><input onclick=""self.location.href='PluginList.asp';"" type=""button"" class=""button"" value=""返回插件管理"" title=""返回插件管理页"" /></p>"
			Response.Write "</form></p>"

		Else
			Response.Write "<p><font color=""red""> × 无法找到插件包. </p>"
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

		If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & SelectedPlugin & "/" & "Plugin.xml") Then

			strXmlFile =BlogPath & "/ZB_USERS/PLUGIN/" & SelectedPlugin & "/" & "Plugin.xml"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else

					Plugin_Author_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					Plugin_Author_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
					Plugin_Author_Email=objXmlFile.documentElement.selectSingleNode("author/email").text

					Plugin_ID=objXmlFile.documentElement.selectSingleNode("id").text
					Plugin_Name=objXmlFile.documentElement.selectSingleNode("name").text
					Plugin_URL=objXmlFile.documentElement.selectSingleNode("url").text
					Plugin_PubDate=objXmlFile.documentElement.selectSingleNode("pubdate").text
					Plugin_Note=objXmlFile.documentElement.selectSingleNode("note").text

					Plugin_Adapted=objXmlFile.documentElement.selectSingleNode("adapted").text
					Plugin_Version=objXmlFile.documentElement.selectSingleNode("version").text
					Plugin_Modified=objXmlFile.documentElement.selectSingleNode("modified").text

					Plugin_Author_Name=TransferHTML(Plugin_Author_Name,"[html-format]")
					Plugin_Name=TransferHTML(Plugin_Name,"[html-format]")


				End If
			Set objXmlFile=Nothing
			End If

			Response.Write "<form id=""edit"" name=""edit"" method=""post"">"

			Response.Write "<p><b>请指定 ZPI 文件中所包含的插件信息, 仅用于基于 Z-blog 1.8 的插件. <a href=""help.asp#aboutzpi"">[什么是 ZPI 文件?]</a></b></p><hr />"

			Response.Write "<p>※插件ID: (插件ID应为插件信息文档中的ID, 此处不可修改.)</p><p><input name=""PluginID"" style=""width:99%"" type=""text"" value="""&Plugin_ID&""" readonly /></p><p></p>"

			Response.Write "<p>※插件名称:</p><p><input name=""PluginName"" style=""width:99%"" type=""text"" value="""&Plugin_Name&""" /></p><p></p>"

			Response.Write "<p>※插件的发布页面: (强列建议填写, 以方便使用者在安装插件前能看到作者的发布页面, 从而获得更多的发布信息.)</p><p><input name=""PluginURL"" style=""width:99%"" type=""text"" value="""&Plugin_URL&""" /></p><p></p>"

			Response.Write "<p>※插件简介 (可以使用 &lt;br /&gt; 换行, 可以使用 html 标签):</p><p><textarea name=""PluginNote"" style=""width:99%"" rows=""5"">"&Plugin_Note&"</textarea></p><p></p>"

			Response.Write "<p><b>以下信息从插件信息文档 Plugin.xml 中读取且必须与之保持一致, 此处不可修改. <a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(SelectedPlugin) & """ title=""编辑插件信息文档-Plugin.xml"">[编辑插件信息文档]</a></b></p><br />"

			Response.Write "<p>※插件适用的 Z-Blog 版本: (写法应为: ""Z-Blog 1.8 Spirit"")</p><p><input name=""PluginAdapted"" style=""width:99%"" type=""text"" value="""&Plugin_Adapted&""" readonly /></p><p></p>"

			Response.Write "<p>※插件的修订版本号:</p><p><input name=""PluginVersion"" style=""width:99%"" type=""text"" value="""&Plugin_Version&""" readonly /></p><p></p>"

			Response.Write "<p>※插件的发布日期: (日期标准格式应为:"""&Date()&""")</p><p><input name=""PluginPubDate"" style=""width:99%"" type=""text"" value="""&Plugin_PubDate&""" readonly /></p><p></p>"

			Response.Write "<p>※插件的最后修改日期: (日期标准格式应为:"""&Date()&""")</p><p><input name=""PluginModified"" style=""width:99%"" type=""text"" value="""&Plugin_Modified&""" readonly /></p><p></p>"

			Response.Write "<p>※插件作者:</p><p><input name=""AuthorName"" style=""width:99%"" type=""text"" value="""&Plugin_Author_Name&""" readonly /></p><p></p>"

			Response.Write "<p>※插件作者主页:</p><p><input name=""AuthorURL"" style=""width:99%"" type=""text"" value="""&Plugin_Author_Url&""" readonly /></p><p></p>"

			Response.Write "<p>※插件作者 Eamil:</p><p><input name=""AuthorEmail"" style=""width:99%"" type=""text"" value="""&Plugin_Author_Email&""" readonly /></p><p></p>"

			Response.Write "<p><b>插件将被按 Z-Blog 插件专用安装包 Version 0.1 打包成 ZPI 文件, 并保存在 TS 插件的 Export 目录下.</b></p><hr />"

			Response.Write "<p><input name=""PubOrBack"" type=""radio"" value=""Pub"" checked=""checked""/> 这次导出是为了<b>发布</b> (导出的文件名必须为 <b>"& SelectedPlugin &".zpi</b>)<br /><input name=""PubOrBack"" type=""radio"" value=""Bak""/> 这次导出是为了<b>备份</b> (导出的文件名为 <b>"& SelectedPlugin &"_Bak(n).zpi</b> 的形式)</p><hr />"


			Response.Write "<p><input type=""submit"" class=""button"" value=""确认信息并打包插件"" id=""btnPost"" onclick='document.getElementById(""edit"").action=""Xml_Pack.asp?act=pack&plugin="& SelectedPlugin &""";' title=""确认信息并打包插件"" /> <input onclick=""self.location.href='PluginList.asp';"" type=""button"" class=""button"" value=""取消并返回插件管理"" title=""取消并返回插件管理页"" />  <input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /></p>"


			Response.Write "</form>"

		Else

			Response.Write "<form id=""edit"" name=""edit"" method=""post"">"
			Response.Write "该插件信息不完全, 不是标准的 Z-Blog 插件, 被打包器拒绝导出! <a href=""javascript:history.back(-1)"" title=""返回上一个页面""><span>[此此返回]</span></a>"
			Response.Write "</form>"

		End If
	Set fso = nothing
	Err.Clear

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
		Xfpath.text = replace(DirPath,ZipPathDir,Pack_PluginDir)

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
					   Xpath.text = replace(PathNameStr,ZipPathDir,Pack_PluginDir)
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

		Set Plugin_ID = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("id"))
			Plugin_ID.Text = Request.Form("PluginID")
		Set Plugin_ID=Nothing

		Set Plugin_Name = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("name"))
			Plugin_Name.Text = Request.Form("PluginName")
		Set Plugin_Name=Nothing

		Set Plugin_URL = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("url"))
			Plugin_URL.Text = Request.Form("PluginURL")
		Set Plugin_URL=Nothing

		Set Plugin_PubDate = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("pubdate"))
			Plugin_PubDate.Text = Request.Form("PluginPubDate")
		Set Plugin_PubDate=Nothing

		Set Plugin_Adapted = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("adapted"))
			Plugin_Adapted.Text = Request.Form("PluginAdapted")
		Set Plugin_Adapted=Nothing

		Set Plugin_Version = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("version"))
			Plugin_Version.Text = Request.Form("PluginVersion")
		Set Plugin_Version=Nothing

		Set Plugin_Modified = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("modified"))
			Plugin_Modified.Text = Request.Form("PluginModified")
		Set Plugin_Modified=Nothing

		Set Plugin_Note = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("note"))
			Plugin_Note.Text = Replace(Replace(Request.Form("PluginNote"),vbCr,""),vbLf,"")
		Set Plugin_Note=Nothing

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
%>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
