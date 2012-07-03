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

BlogTitle="Plugin Sapper"

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
	<div class="Header">Plugin Sapper - 管理您已安装的插件. <a href="help.asp#pluginlist" title="插件管理页帮助">[页面帮助]</a></div>
	<%Call SapperMenu("2")%>
<div id="divMain2">
<%Call GetBlogHint()%>
	<div>
<%
Action=Request.QueryString("act")
NewVersionExists=False

If Action = "plugindel" Then

	SelectedPlugin=Request.QueryString("plugin")
	SelectedPluginName=Request.QueryString("pluginname")

	If CheckPluginState(SelectedPlugin) Then
		Response.Write "<p class=""status-box notice"">您请求的插件正在使用, 无法删除...</p>"
		Response.Write "<script>setTimeout(""self.history.back(1)"",2000);</script>"
		Response.End
	End If

	Dim DelError
	DelError = 0

	If SelectedPlugin<>"" Then
		Response.Write "<p class=""status-box"">正在处理您的请求...</p>"
		Response.Flush

		Response.Write "<p>"
		DelError = DelError + DeleteFolder(BlogPath & "/Plugin/" & SelectedPlugin)
		Response.Write "</p>"
	Else
		Response.Write "<p class=""status-box notice"">请求的参数错误, 正在退出...</p>"
		Response.Flush
		DelError = 13
	End If

	If DelError = 0 Then
		Response.Write "<p><font color=""green""> √ 插件 - " & SelectedPluginName & "  删除成功!</font><p>"
	Else
		Response.Write "<p><font color=""red""> × 插件 - " & SelectedPluginName & "  删除失败! 请手动删除之.</font><p>"
	End If

	Response.Write "<p class=""status-box"">如果您的浏览器没能自动跳转 请 <a href=""PluginList.asp"">[点击这里]</a>.<p>"
	Response.Write "<script>setTimeout(""self.location.href='PluginList.asp'"",1500);</script>"

Else

	Response.Write "<p id=""loading"">正在载入插件列表, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<form id=""edit"" method=""post"" style=""display:none;""><p><a href=""Xml_Restor.asp"" title=""查看主机上保存的 ZPI 文件"">[查看主机上保存的 ZPI 文件]</a> TS 插件的 Export 目录下保存有您已备份或导出的 ZPI 插件文件, 点击可进入查看并对其进行管理操作.</p></form>"

	Response.Write "<p id=""newVersion"" class=""download-box notice"" style=""display:none;""><a href=""Xml_ChkVer.asp"" title=""查看插件的可用更新"">[Plugin Sapper 发现了您安装的某个插件有了新版本, 点此查看现有插件的可用更新]</a>.</p>"

	Dim objXmlFile,strXmlFile
	Dim fso, f, f1, fc, s, t
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "/PLUGIN/")
	Set fc = f.SubFolders

	Dim aryPL
	aryPL=Split(ZC_USING_PLUGIN_LIST,"|")

	For Each s in aryPL

		Plugin_Author_Name=Empty
		Plugin_Author_Url=Empty
		Plugin_Author_Email=Empty

		Plugin_ID=Empty
		Plugin_Name=Empty
		Plugin_URL=Empty
		Plugin_Modified=Empty
		Plugin_Version=Empty
		Plugin_Note=Empty

		strXmlFile =BlogPath & "/ZB_USERS/PLUGIN/" & s & "/" & "Plugin.xml"
		If fso.FileExists(strXmlFile) Then

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
					Plugin_Modified=objXmlFile.documentElement.selectSingleNode("modified").text
					Plugin_Version=objXmlFile.documentElement.selectSingleNode("version").text
					Plugin_Note=objXmlFile.documentElement.selectSingleNode("note").text

					Plugin_Name=TransferHTML(Plugin_Name,"[html-format]")
					Plugin_Note=TransferHTML(Plugin_Note,"[nohtml]")

				End If
			End If
			Set objXmlFile=Nothing

			If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & Plugin_ID & "/" & "verchk.xml") Then
				t="<a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & Plugin_ID) & """ title=""升级插件""><b class=""notice"">发现新版本!</b></a>"
				NewVersionExists=True
			ElseIf fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & Plugin_ID & "/" & "error.log") Then
				t="<b class=""somehow"">不支持在线更新.</b>"
			Else
				t=""
			End If

			Response.Write "<div class=""pluginPanel"">"
			Response.Write "<div class=""listTitle"" onclick='showhidediv("""& Plugin_ID &""");'>"

			Response.Write "<div class=""delicon""><a href=""PluginList.asp?act=plugindel&amp;plugin=" & Server.URLEncode(Plugin_ID) & "&amp;pluginname=" & Server.URLEncode(Plugin_Name) & """ title=""删除此插件"" onclick=""return window.confirm('您将删除此插件的所有文件, 确定吗?');""><font color=""red""><b>×</b></font></a></div>"

			Response.Write "<div class=""epticon""><a href=""Xml_Pack.asp?plugin=" & Server.URLEncode(Plugin_ID) & """ title=""导出此插件""><font color=""green""><b>↑</b></font></a></div>"

			Response.Write "<div class=""edticon""><a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(Plugin_ID) & """ title=""编辑插件信息""><font color=""teal""><b>√</b></font></a></div>"

			Response.Write "<div class=""inficon""><a href=""PluginDetail.asp?plugin=" & Server.URLEncode(Plugin_ID) & "&amp;pluginname=" & Server.URLEncode(Plugin_Name) & """ title=""查看插件信息""><font color=""blue""><b>i</b></font></a></div>"

			Response.Write "<div class=""updicon""><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & Plugin_ID) & """ title=""升级修复插件""><font color=""Gray""><b>↓</b></font></a></div>"

			Response.Write "<div class=""updinfo"">"& Plugin_Name &" Ver "& Plugin_Version &" <font color=""Green"">(启用中)</font> "& t &"</div>"
			Response.Write "</div>"


			Response.Write "<p><table width=""98%"" id="""& Plugin_ID &""" style=""display:none;"">"

			Response.Write "<tr>"

			Response.Write "<td width=""32"" align=""right"">ID:</td><td>"
			Response.Write "" & Plugin_ID & ""
			Response.Write "</td>"

			Response.Write "<td width=""32"" align=""right"">名称:</td><td>"
			If Plugin_URL=Empty Then
				Response.Write "" & Plugin_Name & ""
			Else
				Response.Write "<a href=""" & Plugin_URL & """ target=""_blank"" title=""插件发布地址"">" & Plugin_Name & "</a>"
			End If
			Response.Write "</td>"

			Response.Write "<td width=""32"" align=""right"">作者:</td><td>"
			If Plugin_Author_Url=Empty Then
				Response.Write "" & Plugin_Author_Name & ""
			Else
				Response.Write "<a href=""" & Plugin_Author_Url & """ target=""_blank"" title=""作者主页"">" & Plugin_Author_Name & "</a>"
			End If
			If Not Plugin_Author_Email=Empty Then Response.Write " (" & Plugin_Author_Email & ")"
			Response.Write "</td>"

			Response.Write "<td width=""64"" align=""right"">最后发布:</td><td width=""80"">" & Plugin_Modified & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""right"">简介:</td><td colspan=7>" & Plugin_Note & "</td>"
			Response.Write "</tr>"
			Response.Write "</table></p>"

			Response.Write "</div>"

		End If
	Next

	For Each f1 in fc

		Plugin_Author_Name=Empty
		Plugin_Author_Url=Empty
		Plugin_Author_Email=Empty

		Plugin_ID=Empty
		Plugin_Name=Empty
		Plugin_URL=Empty
		Plugin_Modified=Empty
		Plugin_Version=Empty
		Plugin_Note=Empty


		If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "Plugin.xml") Then

			strXmlFile =BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "Plugin.xml"

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

					'PluginID=f1.name
					Plugin_ID=objXmlFile.documentElement.selectSingleNode("id").text
					Plugin_Name=objXmlFile.documentElement.selectSingleNode("name").text
					Plugin_URL=objXmlFile.documentElement.selectSingleNode("url").text
					Plugin_Modified=objXmlFile.documentElement.selectSingleNode("modified").text
					Plugin_Version=objXmlFile.documentElement.selectSingleNode("version").text
					Plugin_Note=objXmlFile.documentElement.selectSingleNode("note").text

					Plugin_Name=TransferHTML(Plugin_Name,"[html-format]")
					Plugin_Note=TransferHTML(Plugin_Note,"[nohtml]")

				End If
			End If
			Set objXmlFile=Nothing

			If CheckPluginState(Plugin_ID) Then
			Else

			If fso.FileExists(BlogPath & "/PLUGIN/" & Plugin_ID & "/" & "verchk.xml") Then
				t="<a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & Plugin_ID) & """ title=""升级插件""><b class=""notice"">发现新版本!</b></a>"
				NewVersionExists=True
			ElseIf fso.FileExists(BlogPath & "/PLUGIN/" & Plugin_ID & "/" & "error.log") Then
				t="<b class=""somehow"">不支持在线更新.</b>"
			Else
				t=""
			End If

			Response.Write "<div class=""pluginPanel pluginPanelAlt"">"
			Response.Write "<div class=""listTitle listTitleAlt"" onclick='showhidediv("""& Plugin_ID &""");'>"

			Response.Write "<div class=""delicon""><a href=""PluginList.asp?act=plugindel&amp;plugin=" & Server.URLEncode(f1.name) & "&amp;pluginname=" & Server.URLEncode(Plugin_Name) & """ title=""删除此插件"" onclick=""return window.confirm('您将删除此插件的所有文件, 确定吗?');""><font color=""red""><b>×</b></font></a></div>"

			Response.Write "<div class=""epticon""><a href=""Xml_Pack.asp?plugin=" & Server.URLEncode(f1.name) & """ title=""导出此插件""><font color=""green""><b>↑</b></font></a></div>"

			Response.Write "<div class=""edticon""><a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(f1.name) & """ title=""编辑插件信息""><font color=""teal""><b>√</b></font></a></div>"

			Response.Write "<div class=""inficon""><a href=""PluginDetail.asp?plugin=" & Server.URLEncode(f1.name) & "&amp;pluginname=" & Server.URLEncode(Plugin_Name) & """ title=""查看插件信息""><font color=""blue""><b>i</b></font></a></div>"

			Response.Write "<div class=""updicon""><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & Plugin_ID) & """ title=""升级修复插件""><font color=""Gray""><b>↓</b></font></a></div>"

			If UCase(Plugin_ID)<>UCase(f1.name) Then
				Response.Write "<div>该插件ID错误, 请 <a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(f1.name) & """ title=""编辑插件信息""><font color=""red""><b>[重新编辑插件信息]</b></font></a>.</div>"
			Else
				Response.Write "<div class=""updinfo"">"& Plugin_Name &" Ver "& Plugin_Version &" <font color=""Orange"">(停用中)</font> "& t &"</div>"
			End If

			Response.Write "</div>"


			Response.Write "<p><table width=""98%"" id="""& Plugin_ID &""" style=""display:none;"">"

			Response.Write "<tr>"

			Response.Write "<td width=""32"" align=""right"">ID:</td><td>"
			Response.Write "" & Plugin_ID & ""
			Response.Write "</td>"

			Response.Write "<td width=""32"" align=""right"">名称:</td><td>"
			If Plugin_URL=Empty Then
				Response.Write "" & Plugin_Name & ""
			Else
				Response.Write "<a href=""" & Plugin_URL & """ target=""_blank"" title=""插件发布地址"">" & Plugin_Name & "</a>"
			End If
			Response.Write "</td>"

			Response.Write "<td width=""32"" align=""right"">作者:</td><td>"
			If Plugin_Author_Url=Empty Then
				Response.Write "" & Plugin_Author_Name & ""
			Else
				Response.Write "<a href=""" & Plugin_Author_Url & """ target=""_blank"" title=""作者主页"">" & Plugin_Author_Name & "</a>"
			End If
			If Not Plugin_Author_Email=Empty Then Response.Write " (" & Plugin_Author_Email & ")"
			Response.Write "</td>"

			Response.Write "<td width=""64"" align=""right"">最后发布:</td><td>" & Plugin_Modified & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td align=""right"">简介:</td><td colspan=7>" & Plugin_Note & "</td>"
			Response.Write "</tr>"
			Response.Write "</table></p>"

			Response.Write "</div>"

			End If

		End If

	Next

	For Each f1 in fc
		If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "Plugin.xml") Then
		Else

			Plugin_ID=f1.name

			Response.Write "<div class=""pluginPanel"" style=""background-color:#FFFFFF;"">"
			Response.Write "<div class=""listTitle"" style=""border-bottom:1px dotted #BDD3EF;background:#EDEDED;"">"

			Response.Write "<div class=""delicon""><a href=""PluginList.asp?act=plugindel&amp;plugin=" & Server.URLEncode(f1.name) & "&amp;pluginname=" & Server.URLEncode(f1.name) & """ title=""删除此插件"" onclick=""return window.confirm('您将删除此插件的所有文件, 确定吗?');""><font color=""red""><b>×</b></font></a></div>"

			Response.Write "<div class=""epticon""><a href=""Xml_Pack.asp?plugin=" & Server.URLEncode(f1.name) & """ title=""导出此插件""><font color=""green""><b>↑</b></font></a></div>"

			Response.Write "<div class=""edticon""><a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(f1.name) & """ title=""编辑插件信息""><font color=""teal""><b>√</b></font></a></div>"

			Response.Write "<div class=""inficon""><a href=""PluginDetail.asp?plugin=" & Server.URLEncode(f1.name) & "&amp;pluginname=" & Server.URLEncode(f1.name) & """ title=""查看插件信息""><font color=""blue""><b>i</b></font></a></div>"

			Response.Write "<div class=""updicon""><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & Plugin_ID) & """ title=""升级修复插件""><font color=""Gray""><b>↓</b></font></a></div>"

			Response.Write "<div>ID: "& Plugin_ID &"</div>"

			Response.Write "</div>"

			Response.Write "<p>该插件信息不完全, 并不是完整的 Z-Blog 插件.</p>"
			Response.Write "</div>"

		End If

	Next
	Set fso = nothing
	Err.Clear
%>
<!-- 		<div class="PluginPanel" style="background-color:#FFFFFF;">
		<p><a href="Xml_Upload.asp" title="导入本地的 ZPI 文件"><img src="Images/import.png" alt="ScreenShot" width="200" height="160" /></a></p>
			<p><b>从本地导入 ZPI 文件:</b><br />	<form border="1" name="edit" method="post" enctype="multipart/form-data" action="XML_Upload.asp?act=FileUpload"><p>选择插件安装包文件,TS 将从该文件导入插件并安装到 Plugin 目录下: </p><p><input type="file" id="edtFileLoad" name="edtFileLoad" size="15"></p><p><input type="submit" class="button" value="提交" name="B1" onclick="return window.confirm('确定导入该插件数据包??');" /> <input class="button" type="reset" value="重置" name="B2" /></p></form></p>
		</div> -->

		<hr style="clear:both;"/><p><form name="edit" method="get" action="#"  class="status-box">
			<p><input onclick="window.scrollTo(0,0);" type="button" class="button" value="TOP" title="返回页面顶部" /> <input onclick="self.location.href='Xml_ChkVer.asp?act=check&n=0';" type="button" class="button" value="查找更新" title="开始查找插件的可用更新" /></p>
		</form></p>
<%
	If NewVersionExists Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('newVersion').style.display = 'block';</script>"
	End If
	Response.Flush

	Dim FileList,l,c
	FileList=LoadIncludeFiles("ZB_USERS/PLUGIN/PluginSapper/Export/")

	For Each l In FileList
		c=c & l
	Next

	If (InStr(LCase(c),".xml")>0) Or (InStr(LCase(c),".zpi")>0) Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('edit').style.display = 'block';</script>"
	End If

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"

End If
%>
	</div>
</div>
</div>
<SCRIPT type="text/javascript">
function showhidediv(id){
	try{
		if(document.getElementById(id)){
		if(document.getElementById(id).style.display=='none'){
			document.getElementById(id).style.display='block';
		}else{
			document.getElementById(id).style.display='none';
		}
		}
	}catch(e){}
} 
</SCRIPT>
</body>
</html>
<%
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>