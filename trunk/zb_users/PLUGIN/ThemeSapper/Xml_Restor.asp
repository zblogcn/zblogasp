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
<% Server.ScriptTimeout=99999999 %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../function/c_function.asp" -->
<!-- #include file="../../function/c_system_lib.asp" -->
<!-- #include file="../../function/c_system_base.asp" -->
<!-- #include file="../../function/c_system_plugin.asp" -->
<!-- #include file="c_sapper.asp" -->
<!-- #include file="../p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("ThemeSapper")=False Then Call ShowError(48)

BlogTitle = "从本地导入主题"

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
	<div class="Header">Theme Sapper - 管理保存在主机上的 ZTI 文件. <a href="help.asp#restorzti" title="如何管理主机上的 ZTI 文件">[页面帮助]</a></div>
	<%Call SapperMenu("4")%>
<div id="divMain2">
<%
'*********************************************************
%>
<!--以下是显示页面-->
<%
Action=Request.QueryString("act")
SelectedTheme=Request.QueryString("theme")
SelectedThemeName=Request.QueryString("themename")
If SelectedThemeName = "" Then SelectedThemeName = SelectedTheme

If Action="" Then
Call GetBlogHint()

Response.Write "<div>"

	Response.Write "<p id=""loading"">正在载入中, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""1"" cellpadding=""1"" class=""ztiMng"">"

	Dim FileList,l,c
	FileList=LoadIncludeFiles("PLUGIN/ThemeSapper/Export/")

	For Each l In FileList
		c=c & l
	Next

	If (InStr(LCase(c),".xml")>0) Or (InStr(LCase(c),".zti")>0) Then
		Response.Write "<tr><td colspan=5 width='100%'>下面列出了您主机上的保存的 ZTI 主题安装包文件, 您可以下载, 删除这些 ZTI 文件, 或从这些 ZTI 文件恢复主题.</td></tr>"
	Else
		Response.Write "<tr><td colspan=5 width='100%'>对不起, 您的主机上没有保存任何 ZTI 文件! (即 TS 插件中的 Export 目录为空!)</td></tr>"
	End If

	Dim objXmlFile,strXmlFile
	Dim fso, f, f1, fc, s
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "PLUGIN/ThemeSapper/Export/")
	Set fc = f.Files
	For Each f1 in fc

	If GetFileExt(f1.name)="zti" Or GetFileExt(f1.name)="xml" Then

	Response.Write "<tr>"

		strXmlFile =BlogPath & "PLUGIN/ThemeSapper/Export/" & f1.name

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
			ThemeUrl=objXmlFile.documentElement.selectSingleNode("url").text
			ThemeNote=objXmlFile.documentElement.selectSingleNode("note").text
			ThemePubDate=objXmlFile.documentElement.selectSingleNode("pubdate").text
			ThemeModified=objXmlFile.documentElement.selectSingleNode("modified").text

			ThemeNote=TransferHTML(ThemeNote,"[nohtml]")
			If Len(ThemeNote)>42 then ThemeNote=Left(ThemeNote,42-3) & "..."

			End If
		End If
		Set objXmlFile=Nothing


		Response.Write "<td>"& f1.name & "</td>"

		Response.Write "<td align='center'><span class=""rsticon""><a href=""Xml_Restor.asp?act=Restore&amp;id="& Server.URLEncode(ThemeID) &"&amp;theme=" & Server.URLEncode(f1.name) & "&amp;themename=" & Server.URLEncode(ThemeName) & """ title=""从此 ZTI 文件恢复主题到 Blog"">←</a></span></td>"

		If InStr(UCase(Request.ServerVariables("HTTP_USER_AGENT")),UCase("Opera"))>0 Then '如果是Opera浏览器
			Response.Write "<td align='center'><span class=""dowicon""><a href=""Export/"& f1.name & """ title=""右键另存为下载此 ZTI 文件"">↓</a></span></td>"
		Else
			Response.Write "<td align='center'><span class=""dowicon""><a href=""Xml_Download.asp?theme="& Server.URLEncode(f1.name) &""" title=""左键点击下载此 ZTI 文件"">↓</a></span></td>"
		End If

		Response.Write "<td align='center'><span class=""delicon""><a href=""Xml_Restor.asp?act=del&amp;theme=" & Server.URLEncode(f1.name) & "&amp;themename=" & Server.URLEncode(ThemeName) & """ onclick=""return window.confirm('确定删除含有 "& ThemeName &" 的主题数据包 "& f1.name &"?');"" title=""删除此 ZTI 文件"">×</a></span></td><td>"

		'Response.Write "<span>ID:" & ThemeID & "</span> | "

		If ThemeUrl=Empty Then
			Response.Write "<span>" & ThemeName & "</span> "
		Else
			Response.Write "<span><a href=""" & ThemeUrl & """ target=""_blank"">" & ThemeName & "</a></span> "
		End If

		If ThemeAuthor_Url=Empty Then
			Response.Write "<span>by " & ThemeAuthor_Name & "</span> "
		Else
			Response.Write "<span>by " & ThemeAuthor_Name & "</span> "
		End If

		Response.Write "<span>on " & ThemeModified & "</span>"

		Response.Write "<span> | " & ThemeNote & "</span>"
		Response.Write "</td>"

	End If

	Response.Write "</tr>"

	Next
	Set fso = nothing
	Err.Clear

	Response.Write "</table>"

	Response.Write "<p><form id=""edit"" name=""edit"" method=""get"" action=""#"">"
	Response.Write "<p><input onclick=""self.location.href='ThemeList.asp';"" type=""button"" class=""button"" value=""返回主题管理"" title=""返回主题管理页"" /> <input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /></p>"
	Response.Write "</form></p>"

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"

Response.Write "</div>"
End If

Dim Install_Error
Install_Error=0

Dim Install_Pack,Install_Path
Install_Pack = BlogPath & "THEMES/Install.zti"
Install_Path = BlogPath & "THEMES/"


'从主机删除
If Action="del" Then

	Dim DelError
	DelError = 0

	If SelectedTheme<>"" Then
		Response.Write "<p class=""status-box"">正在处理您的请求...</p>"
		Response.Flush

		Response.Write "<p>"
		DelError = DelError + DeleteFile(BlogPath & "/PLUGIN/ThemeSapper/Export/"& SelectedTheme)
		Response.Write "</p>"
	Else
		Response.Write "<p class=""status-box notice"">请求的参数错误, 正在退出...</p>"
		Response.Flush
		DelError = 13
	End If

	If DelError = 0 Then
		Response.Write "<p><font color=""green""> √ 主题安装包 - " & SelectedThemeName & "  删除成功!</font><p>"
	Else
		Response.Write "<p><font color=""red""> × 主题安装包 - " & SelectedThemeName & "  删除失败! 请手动删除之.</font><p>"
	End If

	Response.Write "<p class=""status-box""> 如果您的浏览器没能自动跳转, 请 <a href=""Xml_Restor.asp"">[点击这里]</a>.</p>"
	Response.Write "<script>setTimeout(""self.location.href='Xml_Restor.asp'"",1500);</script>"

End If

'从主机恢复
If Action="Restore" Then

	Call Check_Install()

	Response.Write "<p id=""loading"">正在恢复主题, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<p class=""status-box"">正在复制 ZTI 主题安装包文件 "& SelectedThemeName &"...<p>"
	Response.Flush

	Install_Error=Install_Error + CopyFile(BlogPath & "/PLUGIN/ThemeSapper/Export/"& SelectedTheme,Install_Pack)
	Response.Flush

	Call Install_Theme()

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"

End If


'*********************************************************
Sub Check_Install()
On Error Resume Next

	Dim Confirm,Restor_ID,Alert
	Confirm=Request.QueryString("confirm")
	Restor_ID=Request.QueryString("id")
	Alert=False

	If Confirm<>"yes" Then

		Response.Write "<p id=""loading"">正在校验主题, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
		Response.Flush

		Set objXmlVerChk=New ThemeSapper_CheckVersionViaXML
		objXmlVerChk.XmlDataWeb=(LoadFromFile(BlogPath & "/PLUGIN/ThemeSapper/Export/" & SelectedTheme,"utf-8"))
		objXmlVerChk.XmlDataLocal=(LoadFromFile(BlogPath & "/THEMES/"& Restor_ID &"/Theme.xml","utf-8"))

		If LCase(objXmlVerChk.Item_ID_Web)=LCase(objXmlVerChk.Item_ID_Local) Then
			Response.Write "<p class=""status-box"">您已安装了这个主题 <b>("& objXmlVerChk.Item_Name_Local &")</b>, 是否用 ZTI 文件里的主题 <b>("& objXmlVerChk.Item_Name_Web &")</b> <b>完全覆盖</b>已安装的主题?<br/><br/>"

			Response.Write "您当前主题版本为: <b>"& objXmlVerChk.Item_Version_Local &"</b>. 发布日期为: <b>"& objXmlVerChk.Item_PubDate_Local &"</b>. 最后修改日期为: <b>"& objXmlVerChk.Item_Modified_Local &"</b>.<br/>"
			Response.Write "将要覆盖的版本为: <b>"& objXmlVerChk.Item_Version_Web &"</b>. 发布日期为: <b>"& objXmlVerChk.Item_PubDate_Web &"</b>. 最后修改日期为: <b>"& objXmlVerChk.Item_Modified_Web &"</b><br/><br/>"

			If objXmlVerChk.Item_Url_Web<>Empty Then
				Response.Write "<a href="""& objXmlVerChk.Item_Url_Web &""" target=""_blank"" title=""查看主题的发布页面"">点此查看主题的发布信息!</a><br/><br/>"
			End If

			Response.Write objXmlVerChk.OutputResults & "<br/><br/>"

			Response.Write "<a href=""Xml_Restor.asp?confirm=yes&amp;act=Restore&amp;theme=" & Server.URLEncode(SelectedTheme) & "&amp;themename=" & Server.URLEncode(SelectedThemeName) & """ title=""确认安装"">[确认覆盖]</a> 或 <a href=""javascript:history.back(1);"" title=""返回上一页面"">[取消]</a><p>"
			Alert=True
		End If

		Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"

		Set objXmlVerChk=Nothing

		If Alert=True Then Response.End

	End If

End Sub


Sub Install_Theme()
On Error Resume Next

Response.Write "<p class=""status-box"">ZTI 文件 ""THEMES/Install.zti"" 正在解包安装...<p>"
Response.Flush

	Dim objXmlFile
	Dim objNodeList
	Dim objFSO
	Dim objStream
	Dim i,j

	Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		objXmlFile.load(Install_Pack)
		
		If objXmlFile.readyState<>4 Then
			Response.Write "<p><font color=""red""> × ZTI 文件未准备就绪, 无法解包.</font></p>"
			Install_Error=Install_Error+1
		Else
			If objXmlFile.parseError.errorCode <> 0 Then
				Response.Write "<p><font color=""red""> × ZTI 文件有错误, 无法解包.</font></p>"
				Install_Error=Install_Error+1
			Else

				Dim Pack_ver,Pack_Type,Pack_For,Pack_ID,Pack_Name
				Pack_Ver = objXmlFile.documentElement.SelectSingleNode("//root").getAttributeNode("version").value
				Pack_Type = objXmlFile.documentElement.selectSingleNode("//root").getAttributeNode("type").value
				Pack_For = objXmlFile.documentElement.selectSingleNode("//root").getAttributeNode("for").value
				Pack_ID = objXmlFile.documentElement.selectSingleNode("id").text
				Pack_Name = objXmlFile.documentElement.selectSingleNode("name").text

				If (CDbl(Pack_Ver) > CDbl(XML_Pack_Ver)) Then
					Response.Write "<p><font color=""red""> × ZTI 文件的 XML 版本为 "& Pack_Ver &", 而你的解包器版本为 "& XML_Pack_Ver &", 请升级您的 ThemeSapper, 安装被中止.</font></p>"
					Install_Error=Install_Error+1
				ElseIf (LCase(Pack_Type) <> LCase(XML_Pack_Type)) Then
					Response.Write "<p><font color=""red""> × 不是 ZTI 文件, 而可能是 "& Pack_Type &", 安装被中止.</font></p>"
					Install_Error=Install_Error+1
				ElseIf (LCase(Pack_For) <> LCase(XML_Pack_Version)) Then
					Response.Write "<p><font color=""red""> × ZTI 文件版本不符合, 该版本可能是 "& Pack_For &", 安装被中止.</font></p>"
					Install_Error=Install_Error+1
				Else

					Response.Write "<blockquote><font color=""Teal"">"

					Set objNodeList = objXmlFile.documentElement.selectNodes("//folder/path")
					Set objFSO = CreateObject("Scripting.FileSystemObject")
						
						j=objNodeList.length-1
						For i=0 To j
							If objFSO.FolderExists(Install_Path & objNodeList(i).text)=False Then
								objFSO.CreateFolder(Install_Path & objNodeList(i).text)
							End If
							Response.Write "创建目录" & objNodeList(i).text & "<br/>"
							Response.Flush
						Next
					Set objFSO = Nothing
					Set objNodeList = Nothing
					Set objNodeList = objXmlFile.documentElement.selectNodes("//file/path")
					
						j=objNodeList.length-1
						For i=0 To j
							Set objStream = CreateObject("ADODB.Stream")
								With objStream
									.Type = 1
									.Open
									.Write objNodeList(i).nextSibling.nodeTypedvalue
									.SaveToFile Install_Path & objNodeList(i).text,2
									Response.Write "释放文件" & objNodeList(i).text & "<br/>"
									Response.Flush
									.Close
								End With
							Set objStream = Nothing
						Next
					Set objNodeList = Nothing
					Response.Write "</font></blockquote>"

				End If

			End If
		End If
			
		Set objXmlFile = Nothing

		If Err.Number<>0 Then Install_Error=Install_Error+1
		Err.Clear

		Response.Write "<p>"
		Install_Error = Install_Error + DeleteFile(BlogPath & "THEMES/" & "Install.zti")
		Response.Write "</p>"

		If Install_Error = 0 Then 
			Response.Write "<p>"
			Install_Error = Install_Error + DeleteFile(BlogPath & "THEMES/" & Pack_ID & "/verchk.xml")
			Response.Write "</p>"
		End If

		Response.Flush


	If Install_Error = 0 Then
		Response.Write "<p class=""status-box""> √ 主题恢复完成. 如果您的浏览器没能自动跳转, 请 <a href=""ThemeDetail.asp?theme=" & Server.URLEncode(Pack_ID) & "&amp;themename=" & Server.URLEncode(Pack_Name) & """>[点击这里]</a>.</p>"
		Response.Write "<script>setTimeout(""self.location.href='ThemeDetail.asp?theme=" & Server.URLEncode(Pack_ID) & "&themename=" & Server.URLEncode(Pack_Name) & "'"",3000);</script>"
	Else
		Response.Write "<p class=""status-box""><font color=""red""> × 主题恢复失败. "
		Response.Write "请 <a href=""javascript:window.location.reload();"" title=""刷新此页""><span>[点此重试]</span></a> 或 <a href=""Xml_Restor.asp"" title=""重新上传""><span>[重新上传]</span></a></font></p>"
	End If

End Sub
'*********************************************************
%>
</div>
<script>

	//斑马线
	var tables=document.getElementsByTagName("table");
	var b=false;
	for (var j = 0; j < tables.length; j++){

		var cells = tables[j].getElementsByTagName("tr");

		cells[0].className="color1";
		for (var i = 1; i < cells.length; i++){
			if(b){
				cells[i].className="color2";
				b=false;
			}
			else{
				cells[i].className="color3";
				b=true;
			};
		};
	}

</script>
</body>
</html>
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>