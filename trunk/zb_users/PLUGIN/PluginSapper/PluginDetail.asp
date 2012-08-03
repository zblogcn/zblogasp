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

SelectedPlugin=Request.QueryString("plugin")
SelectedPluginName=Request.QueryString("pluginname")

If SelectedPluginName = "" Then SelectedPluginName = SelectedPlugin

BlogTitle="Plugin Sapper"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">    <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
	<div class="divHeader">Plugin Sapper - 插件: "<%=SelectedPluginName%>" 的详细信息.</div>
	<%Call SapperMenu("0")%>
<div id="divMain2">

	<div>
<%
Response.Write "<p id=""loading"">正在载入插件信息, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
Response.Flush

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

					Plugin_Type=objXmlFile.documentElement.selectSingleNode("type").text
					Plugin_Path=objXmlFile.documentElement.selectSingleNode("path").text
					Plugin_Include=objXmlFile.documentElement.selectSingleNode("include").text
					Plugin_Level=objXmlFile.documentElement.selectSingleNode("level").text

					Plugin_Author_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
					Plugin_Author_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
					Plugin_Author_Email=objXmlFile.documentElement.selectSingleNode("author/email").text

					Plugin_Adapted=objXmlFile.documentElement.selectSingleNode("adapted").text
					Plugin_Version=objXmlFile.documentElement.selectSingleNode("version").text
					Plugin_PubDate=objXmlFile.documentElement.selectSingleNode("pubdate").text
					Plugin_Modified=objXmlFile.documentElement.selectSingleNode("modified").text

			End If
		End If
		Set objXmlFile=Nothing


		If CheckPluginState(Plugin_ID) Then
			Response.Write "<form id=""edit"" name=""edit"" method=""post"" action=""../../cmd.asp?act=PlugInDisable&amp;name="& Plugin_ID &""">"
		Else
			Response.Write "<form id=""edit"" name=""edit"" method=""post"" action=""../../cmd.asp?act=PlugInActive&amp;name="& Plugin_ID &""">"
		End If
		Response.Write "<div class=""PluginDetail"">"

		If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & Plugin_ID & "/" & "verchk.xml") Then
			Response.Write "<p><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & Plugin_ID) & """ title=""升级插件""><b class=""notice"">发现该插件的新版本!</b></a></p><br />"
		ElseIf fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & Plugin_ID & "/" & "error.log") Then
			Response.Write "<p><b class=""somehow"">该插件不支持在线更新.</b></p><br />"
		End If

		If UCase(Plugin_ID)<>UCase(SelectedPlugin) Then
			Response.Write "<p>该插件ID错误, 请 <a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(SelectedPlugin) & """ title=""编辑插件信息""><font color=""red""><b>[重新编辑插件信息]</b></font></a>.</p><br />"
		Else
			Response.Write "<p><b>插件 ID:</b> " & Plugin_ID & "</p><br />"
		End If

		Response.Write "<p><b>插件名称:</b> " & Plugin_Name & "</p>"
		If Plugin_URL<>Empty Then Response.Write "<p><b>发布地址:</b> <a href=""" & Plugin_URL & """ target=""_blank"" title=""插件的发布地址"">" & Plugin_URL & "</a></p>"
		If PluginAuthor_Url=Empty Then
			Response.Write "<p><b>插件作者:</b> " & Plugin_Author_Name & "</p>"
		Else
			Response.Write "<p><b>插件作者:</b> <a href=""" & Plugin_Author_Url & """ target=""_blank"" title=""作者主页"">" & Plugin_Author_Name & "</a></p>"
		End If
		If Plugin_Author_Email<>Empty Then Response.Write "<p><b>作者邮箱:</b> <a href=""mailto:" & Plugin_Author_Email & """ title=""作者邮箱"">" & PluginAuthor_Email & "</a></p>"
		Response.Write "<p><b>发布日期:</b> " & Plugin_PubDate & "</p>"
		Response.Write "<p><b>插件简介:</b> " & Plugin_Note & "</p><br />"

		Response.Write "<p><b>适用于:</b> " & Plugin_Adapted & "</p>"
		Response.Write "<p><b>插件版本:</b> " & Plugin_Version & "</p>"
		Response.Write "<p><b>修正日期:</b> " & Plugin_Modified & "</p><br />"

		Response.Write "<p><b>插件类型:</b> " & Plugin_Type & "</p>"
		If Plugin_Path<>Empty Then Response.Write "<p><b>管理入口:</b> <a href=""../"& Plugin_ID &"/"& Plugin_Path &""">" & Plugin_Path & "</a></p>"
		Response.Write "<p><b>包含文件:</b> " & Plugin_Include & "</p><br />"
		Response.Write "<p><b>操作权限:</b> " & Plugin_Level & "</p><br />"

		Response.Write "<p><b><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & Plugin_ID) & """ title=""升级修复插件"">[升级修复插件]</a>:</b> 重新下载安装此插件以完成对插件的升级和修复.</p>"

		Response.Write "<p><b><a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(SelectedPlugin) & """ title=""编辑插件信息"">[编辑信息]</a>:</b> 此功能可用于生成或编辑该插件的信息文档 Plugin.xml.</p>"

		Response.Write "<p><b><a href=""Xml_Pack.asp?plugin=" & Server.URLEncode(SelectedPlugin) & """ title=""导出插件为 ZPI 文件"">[导出插件]</a>:</b> 将此插件导出为 ZPI 插件安装包文件, 并保存于 TS 插件中的 Export 目录下.</p>"

		Response.Write "<p><b><a href=""PluginList.asp?act=plugindel&amp;plugin=" & Server.URLEncode(SelectedPlugin) & "&amp;pluginname=" & Server.URLEncode(Plugin_Name) & """ title=""删除此插件"" onclick=""return window.confirm('您将删除此插件的所有文件, 确定吗?');"">[删除插件]</a>:</b> 从 PluginS 目录下删除该插件, 正在使用的插件无法删除.</p>"


		If CheckPluginState(Plugin_ID) Then
			Response.Write "</p><br /><p><input type=""submit"" class=""button"" value=""停用此插件"" id=""btnPost"" title=""停用此插件"" />"
		Else
			Response.Write "</p><br /><p><input type=""submit"" class=""button"" value=""启用此插件"" id=""btnPost"" title=""启用此插件"" />"
		End If
		Response.Write " <input onclick=""self.location.href='PluginList.asp';"" type=""button"" class=""button"" value=""返回插件管理"" title=""返回插件管理页"" /> <input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /></p>"

		Response.Write "</div>"
		Response.Write "</form>"

	Else
			Response.Write "<form id=""edit"" name=""edit"" method=""get"" action=""PluginDetail.asp"">"
			Response.Write "<div class=""PluginDetail"">"

			Response.Write "<p><b>该插件信息不完全, 不是标准的 Z-Blog 插件!</b></p><br />"

			Response.Write "<p><b><a href=""Xml_Install.asp?act=update&amp;url=" & Server.URLEncode(Update_URL & SelectedPlugin) & """ title=""升级修复插件"">[升级修复插件]</a>:</b> 重新下载安装此插件以完成对插件的升级和修复.</p>"

			Response.Write "<p><b><a href=""Xml_Edit.asp?plugin=" & Server.URLEncode(SelectedPlugin) & """ title=""编辑插件信息"">[编辑信息]</a>:</b> 此功能可用于生成或编辑该插件的信息文档 Plugin.xml.</p>"

			Response.Write "<p><b><a href=""Xml_Pack.asp?plugin=" & Server.URLEncode(SelectedPlugin) & """ title=""导出插件为 ZPI 文件"">[导出插件]</a>:</b> 将此插件导出为 ZPI 插件安装包文件, 并保存于 TS 插件中的 Export 目录下.</p>"

			Response.Write "<p><b><a href=""PluginList.asp?act=plugindel&amp;plugin=" & Server.URLEncode(SelectedPlugin) & "&amp;pluginname=" & Server.URLEncode(Plugin_Name) & """ title=""删除此插件"" onclick=""return window.confirm('您将删除此插件的所有文件, 确定吗?');"">[删除插件]</a>:</b> 从 PluginS 目录下删除该插件, 正在使用的插件无法删除.</p><br />"

			Response.Write " <p><input onclick=""self.location.href='PluginList.asp';"" type=""button"" class=""button"" value=""返回插件管理"" title=""返回插件管理页"" /></p>"

			Response.Write "</div>"
			Response.Write "</form>"
			End If

	Set fso = nothing
	Err.Clear

Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"
%>
</div></div><!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

