<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Pre Terminator 及以上版本, 其它版本的Z-blog未知
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    主题管理插件
'// 最后修改：   2008-6-28
'// 最后版本:    1.2
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
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

BlogTitle = "编辑主题信息"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta name="robots" content="noindex,nofollow"/>
	<link rel="stylesheet" rev="stylesheet" href="../../CSS/admin.css" type="text/css" media="screen" />
	<link rel="stylesheet" rev="stylesheet" href="images/style.css" type="text/css" media="screen" />

	<script language="JavaScript" src="../../script/common.js" type="text/javascript"></script>
	<script language="JavaScript" src="../../script/jquery.tabs.pack.js" type="text/javascript"></script>
	<link rel="stylesheet" href="../../CSS/jquery.tabs.css" type="text/css" media="print, projection, screen">
	<!--[if lte IE 7]>
	<link rel="stylesheet" href="../../CSS/jquery.tabs-ie.css" type="text/css" media="projection, screen">
	<![endif]-->
	<title><%=BlogTitle%></title>
</head>
<body>
<script language="javascript">
$(document).ready(function(){
	$("#divMain2").tabs({ fxFade: true, fxSpeed: 'fast' });
});
</script>
<div id="divMain">
	<div class="Header">Theme Sapper - 主题信息编辑 - 编辑主题的 Theme.xml 信息文档. <a href="help.asp#editinfo" title="编辑主题信息页帮助">[页面帮助]</a></div>
	<%Call SapperMenu("0")%>
<%
Action=Request.QueryString("act")
SelectedTheme=Request.QueryString("theme")

Response.Write "<div>"
Response.Flush

If Action="" Then
	Response.Write "<p id=""loading"">正在载入主题信息, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Call EditXMLPackInfo()

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('loading').style.display = 'none';</script>"
End If


If Action="save" Then

	Response.Write "<div id=""divMain2"">"

	Response.Write "<p id=""loading2"">正在写入主题信息, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Dim Pack_Error
	Pack_Error=0

	If SelectedTheme="" Then
		Response.Write "<p><font color=""red""> × 主题的名称为空.</font></p>"
		Pack_Error=Pack_Error+1

	Else
		Response.Write "<p class=""status-box""><font color=""Navy"">正在保存XML...</font><p>"
		Response.Flush

		Dim ZipPathFile
		Dim TS_startime,TS_endtime

		'打包文件目录与生成文件名
		ZipPathFile = BlogPath & "THEMES\" & SelectedTheme & "\Theme.xml"

		'开始打包
		CreateXml(ZipPathFile)
	End If

	If Pack_Error = 0 Then
		Call SetBlogHint(True,Empty,Empty)
		Response.Write "<p class=""status-box""><font color=""Navy""> √ 主题信息保存完成. 如果您的浏览器没能自动跳转, 请 <a href=""ThemeDetail.asp?theme="& Server.URLEncode(SelectedTheme) &""">[点击这里]</a>.</font></p>"
		Response.Write "<script>setTimeout(""self.location.href='ThemeDetail.asp?theme="& Server.URLEncode(SelectedTheme) &"'"",1000);</script>"
	Else
		Call SetBlogHint(False,Empty,Empty)
		Response.Write "<p class=""status-box""><font color=""red""> × 主题信息保存失败. "
		Response.Write "<a href=""javascript:history.back(-1)"" title=""返回上一个页面""><span>[返回]</span></a> 或 <a href=""javascript:window.location.reload();"" title=""返回资源列表页""><span>[重试]</span></a></font></p>"
	End If

	Response.Write "</div>"

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('loading2').style.display = 'none';</script>"
End If
Response.Write "</div>"
%>
</div>
</body>
</html>
<%
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

					'ThemeID=SelectedTheme
					ThemeID=objXmlFile.documentElement.selectSingleNode("id").text
					ThemeName=objXmlFile.documentElement.selectSingleNode("name").text
					ThemeURL=objXmlFile.documentElement.selectSingleNode("url").text
					ThemeNote=objXmlFile.documentElement.selectSingleNode("note").text

					ThemeAuthor_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					ThemeAuthor_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
					ThemeAuthor_Email=objXmlFile.documentElement.selectSingleNode("author/email").text

					ThemeSource_Name=objXmlFile.documentElement.selectSingleNode("source/name").text
					ThemeSource_Url=objXmlFile.documentElement.selectSingleNode("source/url").text
					ThemeSource_Email=objXmlFile.documentElement.selectSingleNode("source/email").text

					ThemePlugin_Name=objXmlFile.documentElement.selectSingleNode("plugin/name").text
					ThemePlugin_Note=objXmlFile.documentElement.selectSingleNode("plugin/note").text
					ThemePlugin_Type=objXmlFile.documentElement.selectSingleNode("plugin/type").text
					ThemePlugin_Path=objXmlFile.documentElement.selectSingleNode("plugin/path").text
					ThemePlugin_Include=objXmlFile.documentElement.selectSingleNode("plugin/include").text
					ThemePlugin_Level=objXmlFile.documentElement.selectSingleNode("plugin/level").text

					ThemeAdapted=objXmlFile.documentElement.selectSingleNode("adapted").text
					ThemeVersion=objXmlFile.documentElement.selectSingleNode("version").text
					ThemePubDate=objXmlFile.documentElement.selectSingleNode("pubdate").text
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

			ThemeID=SelectedTheme
			ThemeName=SelectedTheme
			ThemeURL=Empty
			ThemeNote=Empty

			ThemeSource_Name=Empty
			ThemeSource_Url=Empty
			ThemeSource_Email=Empty

			ThemeAuthor_Name=Empty
			ThemeAuthor_Url=Empty
			ThemeAuthor_Email=Empty

			ThemePlugin_Name=Empty
			ThemePlugin_Note=Empty
			ThemePlugin_Type=Empty
			ThemePlugin_Path=Empty
			ThemePlugin_Include=Empty
			ThemePlugin_Level=Empty

			ThemeAdapted="Z-Blog 1.8"
			ThemeVersion=Empty
			ThemePubDate=Date()
			ThemeModified=Date()

			ThemeDescription=Empty

		End If
	Set fso = nothing
	Err.Clear

	Response.Write "<form method=""post"" action=""Xml_Edit.asp?act=save&theme="& SelectedTheme &""">"

	Response.Write "<div id=""divMain2"">"

	Call GetBlogHint()
%>
<ul>
	<li class="tabs-selected"><a href="#fragment-1"><span>主题信息</span></a></li>
	<li><a href="#fragment-2"><span>作者信息</span></a></li>
	<li><a href="#fragment-3"><span>主题说明</span></a></li>
	<li><a href="#fragment-4"><span>主题自带插件(可选)</span></a></li>
</ul>
<%
	Response.Write "<div class=""tabs-div"" id=""fragment-1"">"

	Response.Write "<p>※主题ID: (主题ID应为主题文件夹名称, 由编辑器自动完成填写, 不可修改.)</p><p><input name=""ThemeID"" style=""width:99%"" type=""text"" value="""&SelectedTheme&""" readonly /></p><p></p>"

	Response.Write "<p>※主题名称:</p><p><input name=""ThemeName"" style=""width:99%"" type=""text"" value="""&ThemeName&""" /></p><p></p>"

	Response.Write "<p>※主题的发布页面: (带 http:// 等协议名的页面地址, 以方便使用者获取更多的主题信息)</p><p><input name=""ThemeURL"" style=""width:99%"" type=""text"" value="""&ThemeURL&""" /></p><p></p>"

	Response.Write "<p>※主题简介 (可以用 &lt;br /&gt; 换行, 可以使用 html 标签):</p><p><textarea name=""ThemeNote"" style=""width:99%"" rows=""5"">"&ThemeNote&"</textarea></p><p></p>"

	Response.Write "<p>※适用的 Z-Blog 版本: (要求写法: ""Z-Blog 1.8 Spirit"")</p><p><input name=""ThemeAdapted"" style=""width:99%"" type=""text"" value="""&ThemeAdapted&""" /></p><p></p>"

	Response.Write "<p><b>以下信息对查找主题可用更新极为重要, 建议在每次修改主题后更新这些信息!</a></b></p><hr />"

	Response.Write "<p>※主题的版本号:</p><p><input name=""ThemeVersion"" style=""width:99%"" type=""text"" value="""&ThemeVersion&""" /></p><p></p>"

	Response.Write "<p>※您的主题的发布日期: (日期标准格式:"""&Date()&""")</p><p><input name=""ThemePubDate"" style=""width:99%"" type=""text"" value="""&ThemePubDate&""" /></p><p></p>"

	Response.Write "<p>※最后修改日期: (日期标准格式:"""&Date()&""")</p><p><input name=""ThemeModified"" style=""width:99%"" type=""text"" value="""&ThemeModified&""" /></p><p></p>"

	Response.Write "</div>"
	Response.Write "<div class=""tabs-div"" id=""fragment-2"">"

	Response.Write "<p>※作者名称:</p><p><input name=""ThemeAuthor_Name"" style=""width:99%"" type=""text"" value="""&ThemeAuthor_Name&""" /></p><p></p>"

	Response.Write "<p>※作者网址:</p><p><input name=""ThemeAuthor_URL"" style=""width:99%"" type=""text"" value="""&ThemeAuthor_Url&""" /></p><p></p>"

	Response.Write "<p>※作者 Email:</p><p><input name=""ThemeAuthor_Email"" style=""width:99%"" type=""text"" value="""&ThemeAuthor_Email&""" /></p><p></p>"


	Response.Write "<p>※源作者名称:</p><p><input name=""ThemeSource_Name"" style=""width:99%"" type=""text"" value="""&ThemeSource_Name&""" /></p><p></p>"

	Response.Write "<p>※源作者网址:</p><p><input name=""ThemeSource_URL"" style=""width:99%"" type=""text"" value="""&ThemeSource_Url&""" /></p><p></p>"

	Response.Write "<p>※源作者 Email:</p><p><input name=""ThemeSource_Email"" style=""width:99%"" type=""text"" value="""&ThemeSource_Email&""" /></p><p></p>"

	Response.Write "</div>"
	Response.Write "<div class=""tabs-div"" id=""fragment-3"">"

	Response.Write "<p>※详细说明 (可应用 HTML 代码, 但不必使用换行标签):</p><p><textarea name=""ThemeDescription"" style=""width:99%"" rows=""25"">"&TransferHTML(ThemeDescription,"[textarea]")&"</textarea></p><p></p>"

	Response.Write "</div>"
	Response.Write "<div class=""tabs-div"" id=""fragment-4"">"

	Response.Write "<p>※插件名称:</p><p><input name=""ThemePlugin_Name"" style=""width:99%"" type=""text"" value="""&ThemePlugin_Name&""" /></p><p></p>"

	Response.Write "<p>※插件简要说明:</p><p><input name=""ThemePlugin_Note"" style=""width:99%"" type=""text"" value="""&ThemePlugin_Note&""" /></p><p></p>"

	Response.Write "<p>※插件类型: (挂上接口为 ""inline"", 挂入接口为 ""outline"".)</p><p><input name=""ThemePlugin_Type"" style=""width:99%"" type=""text"" value="""&ThemePlugin_Type&""" /></p><p></p>"

	Response.Write "<p>※插件路径: (插件首页, 如 ""main.asp"".)</p><p><input name=""ThemePlugin_Path"" style=""width:99%"" type=""text"" value="""&ThemePlugin_Path&""" /></p><p></p>"

	Response.Write "<p>※插件 Include 文件: 如 (""include.asp"".)</p><p><input name=""ThemePlugin_Include"" style=""width:99%"" type=""text"" value="""&ThemePlugin_Include&""" /></p><p></p>"

	Response.Write "<p>※插件权限: ( ""1"" 为管理员, ""2"" 为高级用户. 以此类推...)</p><p><input name=""ThemePlugin_Level"" style=""width:99%"" type=""text"" value="""&ThemePlugin_Level&""" /></p><p></p>"

	Response.Write "</div>"

	Response.Write "<hr /><p><b>修改 ID 为 "&SelectedTheme&" 的主题的信息文档. <a href=""help.asp#themexml"">[关于主题信息文档 (Theme.xml)]</a></b></p>"
	Response.Write "<p><b>这些信息将被 (按主题信息文档规范版本 0.1) 保存为 Theme.xml 文件, 该文件将位于主题目录内. <a href=""help.asp#editinfo"">[如何编辑主题信息]</a></b></p><hr />"
	Response.Write "<p><input type=""submit"" class=""button"" value=""完成编辑并保存信息"" id=""btnPost"" title=""完成编辑并保存信息"" /> <input onclick=""self.location.href='ThemeList.asp';"" type=""button"" class=""button"" value=""取消并返回主题管理"" title=""取消并返回主题管理页"" /> <input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /></p>"


	Response.Write "</form>"
	Response.Write "</div>"

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
		Set Root = XmlDoc.createProcessingInstruction("xml","version='1.0' encoding='utf-8' standalone='yes'")
		XmlDoc.appendChild(Root)
		Set xRoot = XmlDoc.appendChild(XmlDoc.CreateElement("theme"))
			xRoot.setAttribute "version",XML_Pack_Ver
		Set xRoot = Nothing

		'写入文件信息

		Dim ThemeAuthor,ThemeSource,ThemePlugin
		Dim XMLcdata

		Set ThemeID = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("id"))
			ThemeID.Text = SelectedTheme
		Set ThemeID=Nothing

		Set ThemeName = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("name"))
			ThemeName.Text = Request.Form("ThemeName")
		Set ThemeName=Nothing

		Set ThemeURL = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("url"))
			ThemeURL.Text = Request.Form("ThemeURL")
		Set ThemeURL=Nothing

		Set ThemeNote = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("note"))
			ThemeNote.Text = Replace(Replace(Request.Form("ThemeNote"),vbCr,""),vbLf,"")
		Set ThemeNote=Nothing


		Set ThemeAuthor = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("author"))

			Set ThemeAuthor_Name = ThemeAuthor.AppendChild(XmlDoc.CreateElement("name"))
				ThemeAuthor_Name.Text = Request.Form("ThemeAuthor_Name")
			Set ThemeAuthor_Name=Nothing

			Set ThemeAuthor_URL = ThemeAuthor.AppendChild(XmlDoc.CreateElement("url"))
				ThemeAuthor_URL.Text = Request.Form("ThemeAuthor_URL")
			Set ThemeAuthor_URL=Nothing

			Set ThemeAuthor_Email = ThemeAuthor.AppendChild(XmlDoc.CreateElement("email"))
				ThemeAuthor_Email.Text = Request.Form("ThemeAuthor_Email")
			Set ThemeAuthor_Email=Nothing

		Set ThemeAuthor=Nothing


		Set ThemeSource = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("source"))

			Set ThemeSource_Name = ThemeSource.AppendChild(XmlDoc.CreateElement("name"))
				ThemeSource_Name.Text = Request.Form("ThemeSource_Name")
			Set ThemeSource_Name=Nothing

			Set ThemeSource_URL = ThemeSource.AppendChild(XmlDoc.CreateElement("url"))
				ThemeSource_URL.Text = Request.Form("ThemeSource_URL")
			Set ThemeSource_URL=Nothing

			Set ThemeSource_Email = ThemeSource.AppendChild(XmlDoc.CreateElement("email"))
				ThemeSource_Email.Text = Request.Form("ThemeSource_Email")
			Set ThemeSource_Email=Nothing

		Set ThemeSource=Nothing

		Set ThemePlugin = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("plugin"))

			Set ThemePlugin_Name = ThemePlugin.AppendChild(XmlDoc.CreateElement("name"))
				ThemePlugin_Name.Text = Request.Form("ThemePlugin_Name")
			Set ThemePlugin_Name=Nothing

			Set ThemePlugin_Note = ThemePlugin.AppendChild(XmlDoc.CreateElement("note"))
				ThemePlugin_Note.Text = Request.Form("ThemePlugin_Note")
			Set ThemePlugin_Note=Nothing

			Set ThemePlugin_Type = ThemePlugin.AppendChild(XmlDoc.CreateElement("type"))
				ThemePlugin_Type.Text = Request.Form("ThemePlugin_Type")
			Set ThemePlugin_Type=Nothing

			Set ThemePlugin_Path = ThemePlugin.AppendChild(XmlDoc.CreateElement("path"))
				ThemePlugin_Path.Text = Request.Form("ThemePlugin_Path")
			Set ThemePlugin_Path=Nothing

			Dim CThemePlugin_Include
			Set ThemePlugin_Include = ThemePlugin.AppendChild(XmlDoc.CreateElement("include"))
				'Set XMLcdata = XmlDoc.createNode("cdatasection", "","")
				'	XMLcdata.NodeValue = Request.Form("ThemePlugin_Include")
				'Set CThemePlugin_Include = ThemePlugin_Include.AppendChild(XMLcdata)
				'Set CThemePlugin_Include = Nothing
				'Set XMLcdata = Nothing
				ThemePlugin_Include.Text = Request.Form("ThemePlugin_Include")
			Set ThemePlugin_Include=Nothing

			Set ThemePlugin_Level = ThemePlugin.AppendChild(XmlDoc.CreateElement("level"))
				ThemePlugin_Level.Text = Request.Form("ThemePlugin_Level")
			Set ThemePlugin_Level=Nothing

		Set ThemePlugin=Nothing


		Set ThemeAdapted = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("adapted"))
			ThemeAdapted.Text = Request.Form("ThemeAdapted")
		Set ThemeAdapted=Nothing

		Set ThemeVersion = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("version"))
			ThemeVersion.Text = Request.Form("ThemeVersion")
		Set ThemeVersion=Nothing

		Set ThemePubDate = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("pubdate"))
			ThemePubDate.Text = Request.Form("ThemePubDate")
		Set ThemePubDate=Nothing

		Set ThemeModified = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("modified"))
			ThemeModified.Text = Request.Form("ThemeModified")
		Set ThemeModified=Nothing


		Dim CThemeDescription
		Set ThemeDescription = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("description"))
			Set XMLcdata = XmlDoc.createNode("cdatasection", "","")
				XMLcdata.NodeValue = Request.Form("ThemeDescription")
			Set CThemeDescription = ThemeDescription.AppendChild(XMLcdata)
			Set CThemeDescription = Nothing
			Set XMLcdata = Nothing
		Set ThemeDescription=Nothing



		XmlDoc.Save(FilePath)
		Set Root = Nothing
	Set XmlDoc = Nothing

	If Err.Number<>0 Then Pack_Error=Pack_Error+1
	Err.Clear

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