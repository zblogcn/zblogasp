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
	<div class="Header">Theme Sapper - 管理您已安装的主题. <a href="help.asp#themelist" title="主题管理面板帮助">[页面帮助]</a></div>
	<%Call SapperMenu("2")%>
<div id="divMain2">
<%Call GetBlogHint()%>
	<div>
<%
Action=Request.QueryString("act")
NewVersionExists=False

If Action = "themedel" Then

	SelectedTheme=Request.QueryString("theme")
	SelectedThemeName=Request.QueryString("themename")

	If UCase(SelectedTheme)=Ucase(ZC_BLOG_THEME) Then
		Response.Write "<p class=""status-box notice"">您请求的主题正在使用, 无法删除...</p>"
		Response.Write "<script>setTimeout(""self.history.back(1)"",2000);</script>"
		Response.End
	End If

	Dim DelError
	DelError = 0

	If SelectedTheme<>"" Then
		Response.Write "<p class=""status-box"">正在处理您的请求...</p>"
		Response.Flush

		Response.Write "<p>"
		DelError = DelError + DeleteFolder(BlogPath & "/THEMES/" & SelectedTheme)
		Response.Write "</p>"
	Else
		Response.Write "<p class=""status-box notice"">请求的参数错误, 正在退出...</p>"
		Response.Flush
		DelError = 13
	End If

	If DelError = 0 Then
		Response.Write "<p><font color=""green""> √ 主题 - " & SelectedThemeName & "  删除成功!</font><p>"
	Else
		Response.Write "<p><font color=""red""> × 主题 - " & SelectedThemeName & "  删除失败! 请手动删除之.</font><p>"
	End If

	Response.Write "<p class=""status-box"">如果您的浏览器没能自动跳转 请 <a href=""ThemeList.asp"">[点击这里]</a>.<p>"
	Response.Write "<script>setTimeout(""self.location.href='ThemeList.asp'"",1500);</script>"

Else

	Response.Write "<p id=""loading"">正在载入主题列表, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<form id=""edit"" method=""post"" style=""display:none;""><p><a href=""Xml_Restor.asp"" title=""查看主机上保存的 ZTI 文件"">[查看主机上保存的 ZTI 文件]</a> TS 插件的 Export 目录下保存有您已备份或导出的 ZTI 主题文件, 点击可进入查看并对其进行管理操作.</p></form>"

	Response.Write "<p id=""newVersion"" class=""download-box notice"" style=""display:none;""><a href=""Xml_ChkVer.asp"" title=""查看主题的可用更新"">[Plugin Sapper 发现了您安装的某个主题有了新版本, 点此查看现有主题的可用更新]</a>.</p>"

	Dim objXmlFile,strXmlFile
	Dim fso, f, f1, fc, s, t
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "/THEMES/")
	Set fc = f.SubFolders
	For Each f1 in fc


		ThemeSource_Name=Empty
		ThemeSource_Url=Empty

		ThemeID=Empty
		ThemeName=Empty
		ThemeURL=Empty
		ThemeNote=Empty
		ThemeModified=Empty

		If fso.FileExists(BlogPath & "/THEMES/" & f1.name & "/" & "theme.xml") Then

			strXmlFile =BlogPath & "/THEMES/" & f1.name & "/" & "theme.xml"

			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else
					ThemeAuthor_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					ThemeAuthor_Url=objXmlFile.documentElement.selectSingleNode("author/url").text

					ThemeSource_Name=objXmlFile.documentElement.selectSingleNode("source/name").text
					ThemeSource_Url=objXmlFile.documentElement.selectSingleNode("source/url").text

					If ThemeAuthor_Name=Empty Then
						ThemeAuthor_Name=ThemeSource_Name
						ThemeAuthor_Url=ThemeSource_Url
					End If

					'ThemeID=f1.name
					ThemeID=objXmlFile.documentElement.selectSingleNode("id").text
					ThemeName=objXmlFile.documentElement.selectSingleNode("name").text
					ThemeURL=objXmlFile.documentElement.selectSingleNode("url").text
					ThemePubDate=objXmlFile.documentElement.selectSingleNode("pubdate").text
					ThemeModified=objXmlFile.documentElement.selectSingleNode("modified").text
					ThemeNote=objXmlFile.documentElement.selectSingleNode("note").text

					If ThemeModified=Empty Then
						ThemeModified=ThemePubDate
					End If

					ThemeNote=TransferHTML(ThemeNote,"[nohtml]")
					If Len(ThemeNote)>25 then ThemeNote=Left(ThemeNote,25-7) & "...<a href=""ThemeDetail.asp?theme=" & Server.URLEncode(ThemeID) & """>more</a>"

				End If
			End If
			Set objXmlFile=Nothing

		Else

			ThemeSource_Name="unknown"
			ThemeSource_Url=Empty

			ThemeID=f1.name
			ThemeName=f1.name
			ThemeURL=Empty
			ThemeNote="unknown"
			ThemeModified="unknown"

		End If

		If fso.FileExists(BlogPath & "/THEMES/" & f1.name & "/" & "verchk.xml") Then
			t="<a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & f1.name) & """ title=""升级主题""><b class=""notice"">发现新版本!</b></a>"
			NewVersionExists=True
		ElseIf fso.FileExists(BlogPath & "/THEMES/" & f1.name & "/" & "error.log") Then
			t="<b class=""somehow"">不支持在线更新.</b>"
		Else
			t=""
		End If

		If fso.FileExists(BlogPath & "/THEMES/" & f1.name & "/" & "screenshot.png") Then
			ThemeScreenShot="../../THEMES/" & f1.name & "/" & "screenshot.png"
		Else
			ThemeScreenShot="Images/noscreenshot.png"
		End If


		If UCase(ThemeID)=UCase(ZC_BLOG_THEME) Then
			Response.Write "<div class=""themePanel"">"
		Else
			Response.Write "<div class=""themePanel"" style=""background-color:#FFFFFF;"">"
		End If

		If UCase(ThemeID)<>UCase(f1.name) Then
			Response.Write "<div>该主题ID错误, 请 <a href=""Xml_Edit.asp?theme=" & Server.URLEncode(f1.name) & """ title=""编辑主题信息""><font color=""red""><b>[重新编辑主题信息]</b></font></a>.</div>"
		Else
			Response.Write "<div class=""delicon""><a href=""ThemeList.asp?act=themedel&amp;theme=" & Server.URLEncode(f1.name) & "&amp;themename=" & Server.URLEncode(ThemeName) & """ title=""删除此主题"" onclick=""return window.confirm('您将删除此主题的所有文件, 确定吗?');""><font color=""red""><b>×</b></font></a></div>"

			Response.Write "<div class=""epticon""><a href=""Xml_Pack.asp?theme=" & Server.URLEncode(f1.name) & """ title=""导出此主题""><font color=""green""><b>↑</b></font></a></div>"

			Response.Write "<div class=""edticon""><a href=""Xml_Edit.asp?theme=" & Server.URLEncode(f1.name) & """ title=""编辑主题信息""><font color=""teal""><b>√</b></font></a></div>"

			Response.Write "<div class=""inficon""><a href=""ThemeDetail.asp?theme=" & Server.URLEncode(f1.name) & "&amp;themename=" & Server.URLEncode(ThemeName) & """ title=""查看主题信息""><font color=""blue""><b>i</b></font></a></div>"

			Response.Write "<div class=""updicon""><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & ThemeID) & """ title=""升级修复主题""><font color=""Gray""><b>↓</b></font></a></div>"

			Response.Write "<div class=""updinfo""><span class=""notice"">"& t &"</span></div>"
		End If

		Response.Write "<p><a href=""ThemeDetail.asp?theme=" & Server.URLEncode(f1.name) & "&amp;themename=" & Server.URLEncode(ThemeName) & """><img src=""" & ThemeScreenShot & """ title=""点击查看 " & ThemeName & " 的详细信息!"" alt=""ScreenShot"" width=""200"" height=""160"" /></a></p>"

		Response.Write "<p><table>"

		If ThemeURL=Empty Then
			Response.Write "<tr><td width=""36"" align=""right"">名称:</td><td>" & ThemeName & "</td></tr>"
		Else
			Response.Write "<tr><td width=""36"" align=""right"">名称:</td><td><a href=""" & ThemeURL & """ target=""_blank"" title=""主题发布地址"">" & ThemeName & "</a></td></tr>"
		End If

		If ThemeAuthor_Url=Empty Then
			Response.Write "<tr><td align=""right"">作者:</td><td>" & ThemeAuthor_Name & "</td></tr>"
		Else
			Response.Write "<tr><td align=""right"">作者:</td><td><a href=""" & ThemeAuthor_Url & """ target=""_blank"" title=""作者主页"">" & ThemeAuthor_Name & "</a></td></tr>"
		End If
		Response.Write "<tr><td align=""right"">发布:</td><td>" & ThemeModified & "</td></tr>"
		Response.Write "<tr><td align=""right"">简介:</td><td>" & ThemeNote & "</td></tr>"
		Response.Write "</table></p>"

		Response.Write "</div>"

	Next
	Set fso = nothing
	Err.Clear
%>
<!-- 		<div class="themePanel" style="background-color:#FFFFFF;">
		<p><a href="Xml_Upload.asp" title="导入本地的 ZTI 文件"><img src="Images/import.png" alt="ScreenShot" width="200" height="160" /></a></p>
			<p><b>从本地导入 ZTI 文件:</b><br />	<form border="1" name="edit" method="post" enctype="multipart/form-data" action="XML_Upload.asp?act=FileUpload"><p>选择主题安装包文件,TS 将从该文件导入主题并安装到 THEMES 目录下: </p><p><input type="file" id="edtFileLoad" name="edtFileLoad" size="15"></p><p><input type="submit" class="button" value="提交" name="B1" onclick="return window.confirm('确定导入该主题数据包??');" /> <input class="button" type="reset" value="重置" name="B2" /></p></form></p>
		</div> -->

		<hr style="clear:both;"/><p><form name="edit" method="get" action="#" class="status-box">
			<p><input onclick="window.scrollTo(0,0);" type="button" class="button" value="TOP" title="返回页面顶部" /> <input onclick="self.location.href='Xml_ChkVer.asp?act=check&n=0';" type="button" class="button" value="查找更新" title="开始查找主题的可用更新" /></p>
		</form></p>
<%
	If NewVersionExists Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('newVersion').style.display = 'block';</script>"
	End If
	Response.Flush

	Dim FileList,l,c
	FileList=LoadIncludeFiles("PLUGIN/ThemeSapper/Export/")

	For Each l In FileList
		c=c & l
	Next

	If (InStr(LCase(c),".xml")>0) Or (InStr(LCase(c),".zti")>0) Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('edit').style.display = 'block';</script>"
	End If

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"

End If
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