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

BlogTitle = "编辑插件信息"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta name="robots" content="noindex,nofollow"/>
	<link rel="stylesheet" rev="stylesheet" href="../../../ZB_SYSTEM/CSS/admin.css" type="text/css" media="screen" />
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
	<div class="Header">Plugin Sapper - 插件信息编辑 - 编辑插件的 Plugin.xml 信息文档. <a href="help.asp#editinfo" title="编辑插件信息页帮助">[页面帮助]</a></div>
	<%Call SapperMenu("0")%>
<%
Action=Request.QueryString("act")
SelectedPlugin=Request.QueryString("Plugin")

Response.Write "<div>"
Response.Flush

If Action="" Then
	Response.Write "<p id=""loading"">正在载入插件信息, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Call EditXMLPackInfo()

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('loading').style.display = 'none';</script>"
End If


If Action="save" Then

	Response.Write "<div id=""divMain2"">"

	Response.Write "<p id=""loading2"">正在写入插件信息, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Dim Pack_Error
	Pack_Error=0

	If SelectedPlugin="" Then
		Response.Write "<p><font color=""red""> × 插件的名称为空.</font></p>"
		Pack_Error=Pack_Error+1

	Else
		Response.Write "<p class=""status-box""><font color=""Navy"">正在保存XML...</font><p>"
		Response.Flush

		Dim ZipPathFile
		Dim TS_startime,TS_endtime

		'打包文件目录与生成文件名
		ZipPathFile = BlogPath & "ZB_USERS\PLUGIN\" & SelectedPlugin & "\Plugin.xml"

		'开始打包
		CreateXml(ZipPathFile)
	End If

	If Pack_Error = 0 Then
		Call SetBlogHint(True,Empty,Empty)
		Response.Write "<p class=""status-box""><font color=""Navy""> √ 插件信息保存完成. 如果您的浏览器没能自动跳转, 请 <a href=""PluginDetail.asp?Plugin="& Server.URLEncode(SelectedPlugin) &""">[点击这里]</a>.</font></p>"
		Response.Write "<script>setTimeout(""self.location.href='PluginDetail.asp?Plugin="& Server.URLEncode(SelectedPlugin) &"'"",1000);</script>"
	Else
		Call SetBlogHint(False,Empty,Empty)
		Response.Write "<p class=""status-box""><font color=""red""> × 插件信息保存失败. "
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

		If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & SelectedPlugin & "/" & "Plugin.xml") Then

			strXmlFile =BlogPath & "/ZB_USERS/PLUGIN/" & SelectedPlugin & "/" & "Plugin.xml"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else

					'Plugin_ID=SelectedPlugin
					Plugin_ID=objXmlFile.documentElement.selectSingleNode("id").text
					Plugin_Name=objXmlFile.documentElement.selectSingleNode("name").text
					Plugin_URL=objXmlFile.documentElement.selectSingleNode("url").text
					Plugin_Note=objXmlFile.documentElement.selectSingleNode("note").text

					Plugin_Author_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					Plugin_Author_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
					Plugin_Author_Email=objXmlFile.documentElement.selectSingleNode("author/email").text

					Plugin_Type=objXmlFile.documentElement.selectSingleNode("type").text
					Plugin_Path=objXmlFile.documentElement.selectSingleNode("path").text
					Plugin_Include=objXmlFile.documentElement.selectSingleNode("include").text
					Plugin_Level=objXmlFile.documentElement.selectSingleNode("level").text

					Plugin_Adapted=objXmlFile.documentElement.selectSingleNode("adapted").text
					Plugin_Version=objXmlFile.documentElement.selectSingleNode("version").text
					Plugin_PubDate=objXmlFile.documentElement.selectSingleNode("pubdate").text
					Plugin_Modified=objXmlFile.documentElement.selectSingleNode("modified").text

					Plugin_Name=TransferHTML(Plugin_Name,"[html-format]")
					Plugin_Author_Name=TransferHTML(Plugin_Author_Name,"[html-format]")

				End If
			Set objXmlFile=Nothing
			End If

		Else

			Plugin_ID=SelectedPlugin
			Plugin_Name=SelectedPlugin
			Plugin_URL=Empty
			Plugin_Note=Empty

			Plugin_Author_Name=Empty
			Plugin_Author_Url=Empty
			Plugin_Author_Email=Empty

			Plugin_Type="outline"
			Plugin_Path="main.asp"
			Plugin_Include="include.asp"
			Plugin_Level="1"

			Plugin_Adapted="Z-Blog 1.8"
			Plugin_Version=Empty
			Plugin_PubDate=Date()
			Plugin_Modified=Date()

		End If
	Set fso = nothing
	Err.Clear

	Response.Write "<form method=""post"" action=""Xml_Edit.asp?act=save&Plugin="& SelectedPlugin &""">"

	Response.Write "<div id=""divMain2"">"

	Call GetBlogHint()
%>
<ul>
	<li class="tabs-selected"><a href="#fragment-1"><span>插件信息</span></a></li>
	<li><a href="#fragment-2"><span>作者信息</span></a></li>
	<li><a href="#fragment-3"><span>插件系统信息</span></a></li>
</ul>
<%
	Response.Write "<div class=""tabs-div"" id=""fragment-1"">"

	Response.Write "<p>※插件ID: (插件ID应为插件文件夹名称, 由编辑器自动完成填写, 不可修改.)</p><p><input name=""PluginID"" style=""width:99%"" type=""text"" value="""&SelectedPlugin&""" readonly /></p><p></p>"

	Response.Write "<p>※插件名称:</p><p><input name=""PluginName"" style=""width:99%"" type=""text"" value="""&Plugin_Name&""" /></p><p></p>"

	Response.Write "<p>※插件的发布页面地址: (带 http:// 等协议名的页面地址, 以方便使用者获取更多的插件发布信息)</p><p><input name=""PluginURL"" style=""width:99%"" type=""text"" value="""&Plugin_URL&""" /></p><p></p>"

	Response.Write "<p>※插件简介 (可以使用 &lt;br /&gt; 换行, 可以使用 html 标签):</p><p><textarea name=""PluginNote"" style=""width:99%"" rows=""5"">"&Plugin_Note&"</textarea></p><p></p>"

	Response.Write "<p>※适用的 Z-Blog 版本: (要求写法: ""Z-Blog 1.8 Spirit"")</p><p><input name=""PluginAdapted"" style=""width:99%"" type=""text"" value="""&Plugin_Adapted&""" /></p><p></p>"

	Response.Write "<p><b>以下信息对查找插件可用更新极为重要, 建议在每次升级插件后更新这些信息!</a></b></p><hr />"

	Response.Write "<p>※插件的版本号:</p><p><input name=""PluginVersion"" style=""width:99%"" type=""text"" value="""&Plugin_Version&""" /></p><p></p>"

	Response.Write "<p>※您的插件的发布日期: (日期标准格式:"""&Date()&""")</p><p><input name=""PluginPubDate"" style=""width:99%"" type=""text"" value="""&Plugin_PubDate&""" /></p><p></p>"

	Response.Write "<p>※最后修改日期: (日期标准格式:"""&Date()&""")</p><p><input name=""PluginModified"" style=""width:99%"" type=""text"" value="""&Plugin_Modified&""" /></p><p></p>"

	Response.Write "</div>"
	Response.Write "<div class=""tabs-div"" id=""fragment-2"">"

	Response.Write "<p>※作者名称:</p><p><input name=""PluginAuthor_Name"" style=""width:99%"" type=""text"" value="""&Plugin_Author_Name&""" /></p><p></p>"

	Response.Write "<p>※作者网址:</p><p><input name=""PluginAuthor_URL"" style=""width:99%"" type=""text"" value="""&Plugin_Author_Url&""" /></p><p></p>"

	Response.Write "<p>※作者 Email:</p><p><input name=""PluginAuthor_Email"" style=""width:99%"" type=""text"" value="""&Plugin_Author_Email&""" /></p><p></p>"

	Response.Write "</div>"


	Response.Write "<div class=""tabs-div"" id=""fragment-3"">"

	Response.Write "<p>※插件类型: (挂上接口为 ""inline"", 挂入接口为 ""outline"".)</p><p><input name=""PluginType"" style=""width:99%"" type=""text"" value="""&Plugin_Type&""" /></p><p></p>"

	Response.Write "<p>※插件路径: (插件首页, 如 ""main.asp"".)</p><p><input name=""PluginPath"" style=""width:99%"" type=""text"" value="""&Plugin_Path&""" /></p><p></p>"

	Response.Write "<p>※插件 Include 文件: 如 (""include.asp"".)</p><p><input name=""PluginInclude"" style=""width:99%"" type=""text"" value="""&Plugin_Include&""" /></p><p></p>"

	Response.Write "<p>※插件权限: ( ""1"" 为管理员, ""2"" 为高级用户. 以此类推...)</p><p><input name=""PluginLevel"" style=""width:99%"" type=""text"" value="""&Plugin_Level&""" /></p><p></p>"

	Response.Write "</div>"

	Response.Write "<hr /><p><b>修改 ID 为 "&Plugin_ID&" 的插件的信息文档. <a href=""help.asp#pluginxml"">[关于插件信息文档 (Plugin.xml)]</a></b></p>"
	Response.Write "<p><b>这些信息将被 (按插件信息文档规范版本 0.1) 保存为 Plugin.xml 文件, 该文件将位于插件目录内. <a href=""help.asp#editinfo"">[如何编辑插件信息]</a></b></p><hr />"
	Response.Write "<p><input type=""submit"" class=""button"" value=""完成编辑并保存信息"" id=""btnPost"" title=""完成编辑并保存信息"" /> <input onclick=""self.location.href='PluginList.asp';"" type=""button"" class=""button"" value=""取消并返回插件管理"" title=""取消并返回插件管理页"" /> <input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /></p>"


	Response.Write "</form>"
	Response.Write "</div>"

End Sub


'创建一个空的XML文件，为写入文件作准备
Sub CreateXml(FilePath)
'On Error Resume Next

	'程序开始执行时间
	TS_startime=timer()
	Dim XmlDoc,Root,xRoot
	Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
		XmlDoc.async = False
		XmlDoc.ValidateOnParse=False
		Set Root = XmlDoc.createProcessingInstruction("xml","version='1.0' encoding='utf-8' standalone='yes'")
		XmlDoc.appendChild(Root)
		Set xRoot = XmlDoc.appendChild(XmlDoc.CreateElement("Plugin"))
			xRoot.setAttribute "version",XML_Pack_Ver
		Set xRoot = Nothing

		'写入文件信息

		Dim Plugin_Author
		Dim XMLcdata

		Set Plugin_ID = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("id"))
			Plugin_ID.Text = SelectedPlugin
		Set Plugin_ID=Nothing

		Set Plugin_Name = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("name"))
			Plugin_Name.Text = Request.Form("PluginName")
		Set Plugin_Name=Nothing

		Set Plugin_URL = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("url"))
			Plugin_URL.Text = Request.Form("PluginURL")
		Set Plugin_URL=Nothing

		Set Plugin_Note = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("note"))
			Plugin_Note.Text = Replace(Replace(Request.Form("PluginNote"),vbCr,""),vbLf,"")
		Set Plugin_Note=Nothing


		Set Plugin_Author = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("author"))

			Set Plugin_Author_Name = Plugin_Author.AppendChild(XmlDoc.CreateElement("name"))
				Plugin_Author_Name.Text = Request.Form("PluginAuthor_Name")
			Set Plugin_Author_Name=Nothing

			Set Plugin_Author_URL = Plugin_Author.AppendChild(XmlDoc.CreateElement("url"))
				Plugin_Author_URL.Text = Request.Form("PluginAuthor_URL")
			Set Plugin_Author_URL=Nothing

			Set Plugin_Author_Email = Plugin_Author.AppendChild(XmlDoc.CreateElement("email"))
				Plugin_Author_Email.Text = Request.Form("PluginAuthor_Email")
			Set Plugin_Author_Email=Nothing

		Set Plugin_Author=Nothing


		Set Plugin_Type = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("type"))
			Plugin_Type.Text = Request.Form("PluginType")
		Set Plugin_Type=Nothing

		Set Plugin_Path = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("path"))
			Plugin_Path.Text = Request.Form("PluginPath")
		Set Plugin_Path=Nothing

		Set Plugin_Include = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("include"))
			Plugin_Include.Text = Request.Form("PluginInclude")
		Set Plugin_Include=Nothing

		Set Plugin_Level = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("level"))
			Plugin_Level.Text = Request.Form("PluginLevel")
		Set Plugin_Level=Nothing


		Set Plugin_Adapted = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("adapted"))
			Plugin_Adapted.Text = Request.Form("PluginAdapted")
		Set Plugin_Adapted=Nothing

		Set Plugin_Version = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("version"))
			Plugin_Version.Text = Request.Form("PluginVersion")
		Set Plugin_Version=Nothing

		Set Plugin_PubDate = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("pubdate"))
			Plugin_PubDate.Text = Request.Form("PluginPubDate")
		Set Plugin_PubDate=Nothing

		Set Plugin_Modified = XmlDoc.SelectSingleNode("//Plugin").AppendChild(XmlDoc.CreateElement("modified"))
			Plugin_Modified.Text = Request.Form("PluginModified")
		Set Plugin_Modified=Nothing

		XmlDoc.Save(FilePath)
		Set Root = Nothing
	Set XmlDoc = Nothing

	'If Err.Number<>0 Then Pack_Error=Pack_Error+1
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