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

SelectedTheme=Request.QueryString("theme")
SelectedThemeName=Request.QueryString("themename")

If SelectedThemeName = "" Then SelectedThemeName = SelectedTheme

BlogTitle="Theme Sapper"

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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
	<div class="Header">Theme Sapper - 主题: "<%=SelectedThemeName%>" 的详细信息.</div>
	<%Call SapperMenu("0")%>
<div id="divMain2">
<%Call GetBlogHint()%>
	<div>
<%
Response.Write "<p id=""loading"">正在载入主题信息, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
Response.Flush

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
				ThemeDescription=TransferHTML(ThemeDescription,"[enter]")

			End If
		End If
		Set objXmlFile=Nothing

	Else

		ThemeID=SelectedTheme
		ThemeName=SelectedTheme
		ThemeURL=Empty
		ThemeNote="Nothing"

		ThemeSource_Name="unknown"
		ThemeSource_Url=Empty
		ThemeSource_Email="unknown"

		ThemeAuthor_Name="unknown"
		ThemeAuthor_Url=Empty
		ThemeAuthor_Email="unknown"

		ThemePlugin_Name="Nothing"
		ThemePlugin_Note="Nothing"
		ThemePlugin_Type="Nothing"
		ThemePlugin_Path="Nothing"
		ThemePlugin_Include="Nothing"
		ThemePlugin_Level="Nothing"

		ThemeAdapted="unknown"
		ThemeVersion="unknown"
		ThemePubDate="unknown"
		ThemeModified="unknown"

		ThemeDescription="Nothing"

	End If

	If fso.FileExists(BlogPath & "/THEMES/" & SelectedTheme & "/" & "screenshot.png") Then
		ThemeScreenShot="../../THEMES/" & SelectedTheme & "/" & "screenshot.png"
	Else
		ThemeScreenShot="Images/noscreenshot.png"
	End If

	Response.Write "<form id=""edit"" name=""edit"" method=""post"" action=""../../cmd.asp?act=ThemesSav"">"
	Response.Write "<div class=""themeDetail"">"

	Response.Write "<p><img src=""" & ThemeScreenShot & """ title=""" & ThemeName & """ alt=""ScreenShot"" /></p>"

	If fso.FileExists(BlogPath & "/THEMES/" & ThemeID & "/" & "verchk.xml") Then
		Response.Write "<p><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & ThemeID) & """ title=""升级主题""><b class=""notice"">发现该主题的新版本!</b></a></p><br />"
	ElseIf fso.FileExists(BlogPath & "/THEMES/" & ThemeID & "/" & "error.log") Then
		Response.Write "<p><b class=""somehow"">该主题不支持在线更新.</b></p><br />"
	End If

	If UCase(ThemeID)<>UCase(SelectedTheme) Then
		Response.Write "<p>该主题ID错误, 请 <a href=""Xml_Edit.asp?theme=" & Server.URLEncode(SelectedTheme) & """ title=""编辑主题信息""><font color=""red""><b>[重新编辑主题信息]</b></font></a>.</p><br />"
	Else
		Response.Write "<p><b>主题 ID:</b> " & ThemeID & "</p><br />"
	End If

	Response.Write "<p><b>主题名称:</b> " & ThemeName & "</p>"
	If ThemeURL<>Empty Then Response.Write "<p><b>发布地址:</b> <a href=""" & ThemeURL & """ target=""_blank"" title=""主题的发布地址"">" & ThemeURL & "</a></p>"
	If ThemeAuthor_Url=Empty Then
		Response.Write "<p><b>主题作者:</b> " & ThemeAuthor_Name & "</p>"
	Else
		Response.Write "<p><b>主题作者:</b> <a href=""" & ThemeAuthor_Url & """ target=""_blank"" title=""作者主页"">" & ThemeAuthor_Name & "</a></p>"
	End If
	If ThemeAuthor_Email<>Empty Then Response.Write "<p><b>作者邮箱:</b> <a href=""mailto:" & ThemeAuthor_Email & """ title=""作者邮箱"">" & ThemeAuthor_Email & "</a></p>"
	Response.Write "<p><b>发布日期:</b> " & ThemePubDate & "</p>"
	Response.Write "<p><b>主题简介:</b> " & ThemeNote & "</p><br />"

	Response.Write "<p><b>适用于:</b> " & ThemeAdapted & "</p>"
	Response.Write "<p><b>主题版本:</b> " & ThemeVersion & "</p>"
	Response.Write "<p><b>修正日期:</b> " & ThemeModified & "</p><br />"

	If ThemeSource_Name<>Empty Then
		If ThemeSource_Url=Empty Then
			Response.Write "<p><b>主题源作者:</b> " & ThemeSource_Name & "</p>"
		Else
			Response.Write "<p><b>主题源作者:</b> <a href=""" & ThemeSource_Url & """ target=""_blank"" title=""源作者主页"">" & ThemeSource_Name & "</a></p>"
		End If
		If ThemeSource_Email<>Empty Then Response.Write "<p><b>源作者邮箱:</b> <a href=""mailto:" & ThemeSource_Email & """ title=""源作者邮箱"">" & ThemeSource_Email & "</a></p>"
	End If

	If ThemePlugin_Name<>Empty Then
		Response.Write "<br /><p><b>此主题含有附带插件:</b></p>"
		Response.Write "<p><b>插件名称:</b> " & ThemePlugin_Name & "</p>"
		Response.Write "<p><b>插件简介:</b> " & ThemePlugin_Note & "</p>"
		Response.Write "<p><b>插件类型:</b> " & ThemePlugin_Type & "</p>"
		Response.Write "<p><b>管理主口:</b> " & ThemePlugin_Path & "</p>"
		Response.Write "<p><b>包含文件:</b> " & ThemePlugin_Include & "</p>"
		Response.Write "<p><b>插件权限:</b> " & ThemePlugin_Level & "</p>"
	End If

	If ThemeDescription<>"" Then Response.Write "<hr /><br /><p><b>详细说明:</b></p>" & "<blockquote>" & ThemeDescription & "</blockquote><br />"

	Response.Write "<p><b><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & ThemeID) & """ title=""升级修复主题"">[升级修复主题]</a>:</b> 重新下载安装此主题以完成对主题的升级和修复.</p>"

	Response.Write "<p><b><a href=""Xml_Edit.asp?theme=" & Server.URLEncode(SelectedTheme) & """ title=""编辑主题信息"">[编辑信息]</a>:</b> 此功能可用于生成或编辑该主题的信息文档 Theme.xml.</p>"

	Response.Write "<p><b><a href=""Xml_Pack.asp?theme=" & Server.URLEncode(SelectedTheme) & """ title=""导出主题为 ZTI 文件"">[导出主题]</a>:</b> 将此主题导出为 ZTI 主题安装包文件, 并保存于 TS 插件中的 Export 目录下.</p>"

	Response.Write "<p><b><a href=""ThemeList.asp?act=themedel&amp;theme=" & Server.URLEncode(SelectedTheme) & "&amp;themename=" & Server.URLEncode(ThemeName) & """ title=""删除此主题"" onclick=""return window.confirm('您将删除此主题的所有文件, 确定吗?');"">[删除主题]</a>:</b> 从 THEMES 目录下删除该主题, 正在使用的主题无法删除.</p>"


	Dim i,j
	Dim aryFileList
	Dim Theme_Style_Name

	aryFileList=LoadIncludeFiles("themes" & "/" & SelectedTheme & "/style")

	Response.Write "<br /><hr /><p><input type=""hidden"" name=""edtZC_BLOG_THEME"" value="""& SelectedTheme &""">"

	If IsArray(aryFileList) Then
		j=UBound(aryFileList)
		For i=1 to j
		GetFileExt(f1.name)="zti" Or GetFileExt(f1.name)="xml"
			If (GetFileExt(aryFileList(i))="css") Or (GetFileExt(aryFileList(i))="asp") Then
				Theme_Style_Name=Replace(aryFileList(i),"." & GetFileExt(aryFileList(i)),"")
				Response.Write "<p>"
				If i=1 Then
					Response.Write " <input type=""radio"" checked=""checked"" name=""edtZC_BLOG_CSS"" value="""& Theme_Style_Name &"""> 选择样式 "&aryFileList(i)&"; "
				Else
					Response.Write " <input type=""radio"" name=""edtZC_BLOG_CSS"" value="""& Theme_Style_Name &"""> 选择样式 "&aryFileList(i)&"; "
				End If
				Response.Write "</p>"
			End If
		Next
	End If

	Response.Write "</p><br /><p><input type=""submit"" class=""button"" value=""应用此主题"" id=""btnPost"" title=""应用此主题"" /> <input onclick=""self.location.href='ThemeList.asp';"" type=""button"" class=""button"" value=""返回主题管理"" title=""返回主题管理页"" /> <input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /></p>"

	Response.Write "</div>"
	Response.Write "</form>"

	Set fso = nothing
	Err.Clear

Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"
%>
	</div>
</div>
</div>
</body>
</html>
<%
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>