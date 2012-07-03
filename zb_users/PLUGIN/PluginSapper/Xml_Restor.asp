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
<!-- #include file="../p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("PluginSapper")=False Then Call ShowError(48)

BlogTitle = "从本地导入插件"

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta name="robots" content="noindex,nofollow"/>
	<link rel="stylesheet" rev="stylesheet" href="../../../ZB_SYSTEM/CSS/admin.css" type="text/css" media="screen" />
	<link rel="stylesheet" rev="stylesheet" href="images/style.css" type="text/css" media="screen" />
	<title><%=BlogTitle%></title>
</head>
<body>

<div id="divMain">
	<div class="Header">Plugin Sapper - 管理保存在主机上的 ZPI 文件. <a href="help.asp#restorzpi" title="如何管理主机上的 ZPI 文件">[页面帮助]</a></div>
	<%Call SapperMenu("4")%>
<div id="divMain2">
<%
'*********************************************************
%>
<!--以下是显示页面-->
<%
Action=Request.QueryString("act")
SelectedPlugin=Request.QueryString("plugin")
SelectedPluginName=Request.QueryString("pluginname")
If SelectedPluginName = "" Then SelectedPluginName = SelectedPlugin

If Action="" Then
Call GetBlogHint()

Response.Write "<div>"

	Response.Write "<p id=""loading"">正在载入中, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<table border=""1"" width=""100%"" cellspacing=""1"" cellpadding=""1"" class=""zpiMng"">"

	Dim FileList,l,c
	FileList=LoadIncludeFiles("ZB_USERS/PLUGIN/PluginSapper/Export/")

	For Each l In FileList
		c=c & l
	Next

	If (InStr(LCase(c),".xml")>0) Or (InStr(LCase(c),".zpi")>0) Then
		Response.Write "<tr><td colspan=5 width='100%'>下面列出了您主机上的保存的 ZPI 插件安装包文件, 您可以下载, 删除这些 ZPI 文件, 或从这些 ZPI 文件恢复插件.</td></tr>"
	Else
		Response.Write "<tr><td colspan=5 width='100%'>对不起, 您的主机上没有保存任何 ZPI 文件! (即 TS 插件中的 Export 目录为空!)</td></tr>"
	End If

	Dim objXmlFile,strXmlFile
	Dim fso, f, f1, fc, s
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "ZB_USERS/PLUGIN/PluginSapper/Export/")
	Set fc = f.Files
	For Each f1 in fc

	If GetFileExt(f1.name)="zpi" Or GetFileExt(f1.name)="xml" Then

	Response.Write "<tr>"

		strXmlFile =BlogPath & "ZB_USERS/PLUGIN/PluginSapper/Export/" & f1.name

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
			Plugin_Url=objXmlFile.documentElement.selectSingleNode("url").text
			Plugin_Note=objXmlFile.documentElement.selectSingleNode("note").text
			Plugin_PubDate=objXmlFile.documentElement.selectSingleNode("pubdate").text
			Plugin_Modified=objXmlFile.documentElement.selectSingleNode("modified").text

			Plugin_Name=TransferHTML(Plugin_Name,"[html-format]")
			Plugin_Note=TransferHTML(Plugin_Note,"[nohtml]")
			If Len(Plugin_Note)>42 then Plugin_Note=Left(Plugin_Note,42-3) & "..."

			End If
		End If
		Set objXmlFile=Nothing


		Response.Write "<td>"& f1.name & "</td>"

		Response.Write "<td align='center'><span class=""rsticon""><a href=""Xml_Restor.asp?act=Restore&amp;id="& Server.URLEncode(Plugin_ID) &"&amp;plugin=" & Server.URLEncode(f1.name) & "&amp;pluginname=" & Server.URLEncode(Plugin_Name) & """ title=""从此 ZPI 文件恢复插件到 Blog"">←</a></span></td>"

		If InStr(UCase(Request.ServerVariables("HTTP_USER_AGENT")),UCase("Opera"))>0 Then '如果是Opera浏览器
			Response.Write "<td align='center'><span class=""dowicon""><a href=""Export/"& f1.name & """ title=""右键另存为下载此 ZPI 文件"">↓</a></span></td>"
		Else
			Response.Write "<td align='center'><span class=""dowicon""><a href=""Xml_Download.asp?plugin="& Server.URLEncode(f1.name) &""" title=""左键点击下载此 ZPI 文件"">↓</a></span></td>"
		End If

		Response.Write "<td align='center'><span class=""delicon""><a href=""Xml_Restor.asp?act=del&amp;plugin=" & Server.URLEncode(f1.name) & "&amp;pluginname=" & Server.URLEncode(Plugin_Name) & """ onclick=""return window.confirm('确定删除含有 "& Plugin_Name &" 的插件数据包 "& f1.name &"?');"" title=""删除此 ZPI 文件"">×</a></span></td><td>"

		'Response.Write "<span>ID:" & Plugin_ID & "</span> | "

		If Plugin_Url=Empty Then
			Response.Write "<span>" & Plugin_Name & "</span> "
		Else
			Response.Write "<span><a href=""" & Plugin_Url & """ target=""_blank"">" & Plugin_Name & "</a></span> "
		End If

		If Plugin_Author_Url=Empty Then
			Response.Write "<span>by " & Plugin_Author_Name & "</span> "
		Else
			Response.Write "<span>by " & Plugin_Author_Name & "</span> "
		End If

		Response.Write "<span>on " & Plugin_Modified & "</span>"

		Response.Write "<span> | " & Plugin_Note & "</span>"
		Response.Write "</td>"

	End If

	Response.Write "</tr>"

	Next
	Set fso = nothing
	Err.Clear

	Response.Write "</table>"

	Response.Write "<p><form id=""edit"" name=""edit"" method=""get"" action=""#"">"
	Response.Write "<p><input onclick=""self.location.href='PluginList.asp';"" type=""button"" class=""button"" value=""返回插件管理"" title=""返回插件管理页"" /> <input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /></p>"
	Response.Write "</form></p>"

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"

Response.Write "</div>"
End If

Dim Install_Error
Install_Error=0

Dim Install_Pack,Install_Path
Install_Pack = BlogPath & "ZB_USERS/PLUGIN/Install.zpi"
Install_Path = BlogPath & "ZB_USERS/PLUGIN/"


'从主机删除
If Action="del" Then

	Dim DelError
	DelError = 0

	If SelectedPlugin<>"" Then
		Response.Write "<p class=""status-box"">正在处理您的请求...</p>"
		Response.Flush

		Response.Write "<p>"
		DelError = DelError + DeleteFile(BlogPath & "/ZB_USERS/PLUGIN/PluginSapper/Export/"& SelectedPlugin)
		Response.Write "</p>"
	Else
		Response.Write "<p class=""status-box notice"">请求的参数错误, 正在退出...</p>"
		Response.Flush
		DelError = 13
	End If

	If DelError = 0 Then
		Response.Write "<p><font color=""green""> √ 插件安装包 - " & SelectedPluginName & "  删除成功!</font><p>"
	Else
		Response.Write "<p><font color=""red""> × 插件安装包 - " & SelectedPluginName & "  删除失败! 请手动删除之.</font><p>"
	End If

	Response.Write "<p class=""status-box""> 如果您的浏览器没能自动跳转, 请 <a href=""Xml_Restor.asp"">[点击这里]</a>.</p>"
	Response.Write "<script>setTimeout(""self.location.href='Xml_Restor.asp'"",1500);</script>"

End If

'从主机恢复
If Action="Restore" Then

	Call Check_Install()

	Response.Write "<p id=""loading"">正在恢复插件, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<p class=""status-box"">正在复制 ZPI 插件安装包文件 "& SelectedPluginName &"...<p>"
	Response.Flush

	Install_Error=Install_Error + CopyFile(BlogPath & "/ZB_USERS/PLUGIN/PluginSapper/Export/"& SelectedPlugin,Install_Pack)
	Response.Flush

	Call Install_Plugin()

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

		Response.Write "<p id=""loading"">正在校验插件, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
		Response.Flush

		Set objXmlVerChk=New PluginSapper_CheckVersionViaXML
		objXmlVerChk.XmlDataWeb=(LoadFromFile(BlogPath & "/ZB_USERS/PLUGIN/PluginSapper/Export/" & SelectedPlugin,"utf-8"))
		objXmlVerChk.XmlDataLocal=(LoadFromFile(BlogPath & "/ZB_USERS/PLUGIN/"& Restor_ID &"/plugin.xml","utf-8"))

		If LCase(objXmlVerChk.Item_ID_Web)=LCase(objXmlVerChk.Item_ID_Local) Then
			Response.Write "<p class=""status-box"">您已安装了这个插件 <b>("& objXmlVerChk.Item_Name_Local &")</b>, 是否用 ZPI 文件 <b>("& objXmlVerChk.Item_Name_Web &")</b> 里的插件<b>完全覆盖</b>已安装的插件?<br/><br/>"

			Response.Write "您当前插件版本为: <b>"& objXmlVerChk.Item_Version_Local &"</b>. 发布日期为: <b>"& objXmlVerChk.Item_PubDate_Local &"</b>. 最后修改日期为: <b>"& objXmlVerChk.Item_Modified_Local &"</b>.<br/>"
			Response.Write "将要覆盖的版本为: <b>"& objXmlVerChk.Item_Version_Web &"</b>. 发布日期为: <b>"& objXmlVerChk.Item_PubDate_Web &"</b>. 最后修改日期为: <b>"& objXmlVerChk.Item_Modified_Web &"</b><br/><br/>"

			If objXmlVerChk.Item_Url_Web<>Empty Then
				Response.Write "<a href="""& objXmlVerChk.Item_Url_Web &""" target=""_blank"" title=""查看插件的发布页面"">点此查看插件的发布信息!</a><br/><br/>"
			End If

			Response.Write objXmlVerChk.OutputResults & "<br/><br/>"

			Response.Write "<a href=""Xml_Restor.asp?confirm=yes&amp;act=Restore&amp;plugin=" & Server.URLEncode(SelectedPlugin) & "&amp;pluginname=" & Server.URLEncode(SelectedPluginName) & """ title=""确认安装"">[确认覆盖]</a> 或 <a href=""javascript:history.back(1);"" title=""返回上一页面"">[取消]</a><p>"
			Alert=True
		End If

		Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"

		Set objXmlVerChk=Nothing

		If Alert=True Then Response.End

	End If

End Sub


Sub Install_Plugin()
On Error Resume Next

Response.Write "<p class=""status-box"">ZPI 文件 ""PluginS/Install.zpi"" 正在解包安装...<p>"
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
			Response.Write "<p><font color=""red""> × ZPI 文件未准备就绪, 无法解包.</font></p>"
			Install_Error=Install_Error+1
		Else
			If objXmlFile.parseError.errorCode <> 0 Then
				Response.Write "<p><font color=""red""> × ZPI 文件有错误, 无法解包.</font></p>"
				Install_Error=Install_Error+1
			Else

				Dim Pack_ver,Pack_Type,Pack_For,Pack_ID,Pack_Name
				Pack_Ver = objXmlFile.documentElement.SelectSingleNode("//root").getAttributeNode("version").value
				Pack_Type = objXmlFile.documentElement.selectSingleNode("//root").getAttributeNode("type").value
				Pack_For = objXmlFile.documentElement.selectSingleNode("//root").getAttributeNode("for").value
				Pack_ID = objXmlFile.documentElement.selectSingleNode("id").text
				Pack_Name = objXmlFile.documentElement.selectSingleNode("name").text

				If (CDbl(Pack_Ver) > CDbl(XML_Pack_Ver)) Then
					Response.Write "<p><font color=""red""> × ZPI 文件的 XML 版本为 "& Pack_Ver &", 而你的解包器版本为 "& XML_Pack_Ver &", 请升级您的 PluginSapper, 安装被中止.</font></p>"
					Install_Error=Install_Error+1
				ElseIf (LCase(Pack_Type) <> LCase(XML_Pack_Type)) Then
					Response.Write "<p><font color=""red""> × 不是 ZPI 文件, 而可能是 "& Pack_Type &", 安装被中止.</font></p>"
					Install_Error=Install_Error+1
				ElseIf (LCase(Pack_For) <> LCase(XML_Pack_Version)) Then
					Response.Write "<p><font color=""red""> × ZPI 文件版本不符合, 该版本可能是 "& Pack_For &", 安装被中止.</font></p>"
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
		Install_Error = Install_Error + DeleteFile(BlogPath & "ZB_USERS/PLUGIN/" & "Install.zpi")
		Response.Write "</p>"

		If Install_Error = 0 Then
			Response.Write "<p>"
			Install_Error = Install_Error + DeleteFile(BlogPath & "ZB_USERS/PLUGIN/" & Pack_ID & "/verchk.xml")
			Response.Write "</p>"
		End If

		Response.Flush


	If Install_Error = 0 Then
		Response.Write "<p class=""status-box""> √ 插件恢复完成. 如果您的浏览器没能自动跳转, 请 <a href=""PluginDetail.asp?plugin=" & Server.URLEncode(Pack_ID) & "&amp;pluginname=" & Server.URLEncode(Pack_Name) & """>[点击这里]</a>.</p>"
		Response.Write "<script>setTimeout(""self.location.href='PluginDetail.asp?plugin=" & Server.URLEncode(Pack_ID) & "&pluginname=" & Server.URLEncode(Pack_Name) & "'"",3000);</script>"
	Else
		Response.Write "<p class=""status-box""><font color=""red""> × 插件恢复失败. "
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