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

Action=Request.QueryString("act")

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
	<div class="Header">Plugin Sapper - 查看插件的可用更新. <a href="help.asp#checkupdate" title="查看插件的可用更新帮助">[页面帮助]</a></div>
	<%Call SapperMenu("5")%>
<div id="divMain2">
<%Call GetBlogHint()%>
	<div>
<%
Dim fso, f, f1, fc, s, t, i, n, m

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFolder(BlogPath & "/ZB_USERS/PLUGIN/")
Set fc = f.SubFolders

If Action="" Then
	Response.Write "<p id=""loading"">正在载入页面, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<p class=""status-box"" id=""found"" style=""display:none;"">!! 以下列出了您需要更新的插件:</p>"
	Response.Write "<p class=""status-box"" id=""notfound"" style=""display:none;"">!! 暂时没有找到需要更新的插件.</p>"
	Response.Flush

	m=False

	For Each f1 in fc

		Set objXmlVerChk=New PluginSapper_CheckVersionViaXML

		If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "Plugin.xml") Then

			objXmlVerChk.XmlDataLocal=(LoadFromFile(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/plugin.xml","utf-8"))

			If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "verchk.xml") Then

				objXmlVerChk.XmlDataWeb=(LoadFromFile(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "verchk.xml","utf-8"))

				Response.Write "<p class=""download-box"">"
				Response.Write "插件: <b>"& objXmlVerChk.Item_Name_Local &"</b> <b class=""notice"">发现可用的新版本!</b><br/><br/>"

				Response.Write "当前插件版本为: <b>"& objXmlVerChk.Item_Version_Local &"</b>. 发布日期为: <b>"& objXmlVerChk.Item_PubDate_Local &"</b>. 最后修改日期为: <b>"& objXmlVerChk.Item_Modified_Local &"</b>.<br/>"
				Response.Write "最新可用版本为: <b>"& objXmlVerChk.Item_Version_Web &"</b>. 发布日期为: <b>"& objXmlVerChk.Item_PubDate_Web &"</b>. 最后修改日期为: <b>"& objXmlVerChk.Item_Modified_Web &"</b><br/><br/>"

				If objXmlVerChk.Item_Url_Web<>Empty Then
					Response.Write "<a href="""& objXmlVerChk.Item_Url_Web &""" target=""_blank"" title=""查看插件的发布页面"">[点此查看插件的发布信息!]</a> "
				End If

				Response.Write "<a href=""Xml_Install.asp?act=confirm&amp;url=" & Server.URLEncode(Update_URL & f1.name) & """ title=""升级此插件"">[升级此插件]</a>"

				Response.Write "</p>"
				Response.Flush

				m=True

			End If

		End If

		Set objXmlVerChk=Nothing

	Next

	If m Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('found').style.display = 'block';}catch(e){};</script>"
	Else
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('notfound').style.display = 'block';}catch(e){};</script>"
	End If

	Response.Write "<form name=""edit"" method=""get"" action=""#"" class=""status-box"">"
	Response.Write "<p><input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /> <input onclick=""self.location.href='Xml_ChkVer.asp?act=check&n=0';"" type=""button"" class=""button"" value=""查找更新"" title=""开始查找插件的可用更新"" /> <input onclick=""self.location.href='Xml_ChkVer.asp?act=unsupport';"" type=""button"" class=""button"" value=""列出不支持在线更新的插件"" title=""列出不支持在线更新的插件"" /> <input onclick=""self.location.href='Xml_ChkVer.asp?act=clear';"" type=""button"" class=""button"" value=""清除更新提示"" title=""清除更新提示"" /></p>"
	Response.Write "</form>"

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"
End If


If Action="unsupport" Then
	Response.Write "<p id=""loading"">正在载入页面, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<p class=""status-box"" id=""found"" style=""display:none;"">!! 以下列出了不支持在线更新的插件:</p>"
	Response.Write "<p class=""status-box"" id=""notfound"" style=""display:none;"">!! 暂时没发现不支持在线更新的插件.</p>"
	Response.Flush

	m=False

	For Each f1 in fc

		Set objXmlVerChk=New PluginSapper_CheckVersionViaXML

		If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "Plugin.xml") Then

			objXmlVerChk.XmlDataLocal=(LoadFromFile(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/plugin.xml","utf-8"))

			If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "error.log") Then

				Response.Write "<p class=""download-box"">"
				Response.Write "插件: <b>"& objXmlVerChk.Item_Name_Local &"</b> <b class=""somehow"">可能不支持在线更新!</b><br/><br/>"

				Response.Write "当前插件版本为: <b>"& objXmlVerChk.Item_Version_Local &"</b>. 发布日期为: <b>"& objXmlVerChk.Item_PubDate_Local &"</b>. 最后修改日期为: <b>"& objXmlVerChk.Item_Modified_Local &"</b>.<br/><br/>"

				If objXmlVerChk.Item_Url_Local<>Empty Then
					Response.Write "<a href="""& objXmlVerChk.Item_Url_Local &""" target=""_blank"" title=""查看插件的发布页面"">[点此查看插件的发布信息!]</a> "
				End If

				Response.Write "</p>"
				Response.Flush

				m=True

			End If

		End If

		Set objXmlVerChk=Nothing

	Next

	If m Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('found').style.display = 'block';}catch(e){};</script>"
	Else
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('notfound').style.display = 'block';}catch(e){};</script>"
	End If

	Response.Write "<form name=""edit"" method=""get"" action=""#"" class=""status-box"">"
	Response.Write "<p><input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /> <input onclick=""self.location.href='Xml_ChkVer.asp?act=check&n=0';"" type=""button"" class=""button"" value=""查找更新"" title=""开始查找插件的可用更新"" /> <input onclick=""self.location.href='Xml_ChkVer.asp?act=unsupport';"" type=""button"" class=""button"" value=""列出不支持在线更新的插件"" title=""列出不支持在线更新的插件"" /> <input onclick=""self.location.href='Xml_ChkVer.asp?act=clear';"" type=""button"" class=""button"" value=""清除更新提示"" title=""清除更新提示"" /></p>"
	Response.Write "</form>"

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"
End If


If Action="clear" Then
	Response.Write "<p id=""loading"">正在载入页面, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<p class=""status-box"" id=""found"" style=""display:none;"">!! 已清除以下插件的更新提示:</p>"
	Response.Write "<p class=""status-box"" id=""notfound"" style=""display:none;"">!! 没有更新提示可清除.</p>"
	Response.Flush

	m=False

	For Each f1 in fc

		Set objXmlVerChk=New PluginSapper_CheckVersionViaXML

		If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "Plugin.xml") Then

			objXmlVerChk.XmlDataLocal=(LoadFromFile(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/plugin.xml","utf-8"))

			If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "verchk.xml") Then

				fso.DeleteFile(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "verchk.xml")

				Response.Write "<p class=""status-box"">"
				Response.Write "插件: <b>"& objXmlVerChk.Item_Name_Local &"</b> <b class=""allright"">的新版提示已清除!</b>"
				Response.Write "</p>"
				Response.Flush

				m=True

			End If

			If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "error.log") Then

				fso.DeleteFile(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "error.log")

				Response.Write "<p class=""status-box"">"
				Response.Write "插件: <b>"& objXmlVerChk.Item_Name_Local &"</b> <b class=""allright"">的不支持更新提示已清除!</b>"
				Response.Write "</p>"
				Response.Flush

				m=True

			End If

		End If

		Set objXmlVerChk=Nothing

	Next

	If m Then
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('found').style.display = 'block';}catch(e){};</script>"
	Else
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('notfound').style.display = 'block';}catch(e){};</script>"
	End If

	Response.Write "<form name=""edit"" method=""get"" action=""#"" class=""status-box"">"
	Response.Write "<p><input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /> <input onclick=""self.location.href='Xml_ChkVer.asp?act=check&n=0';"" type=""button"" class=""button"" value=""查找更新"" title=""开始查找插件的可用更新"" /> <input onclick=""self.location.href='Xml_ChkVer.asp?act=unsupport';"" type=""button"" class=""button"" value=""列出不支持在线更新的插件"" title=""列出不支持在线更新的插件"" /> <input onclick=""self.location.href='Xml_ChkVer.asp?act=clear';"" type=""button"" class=""button"" value=""清除更新提示"" title=""清除更新提示"" /></p>"
	Response.Write "</form>"

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"
End If

If Action="check" Then
	Response.Write "<p id=""loading2"">正在查找更新, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<p class=""status-box"">!! 查找更新的过程会消耗一些时间, 时间长短会因您安装的插件数量而异, 请耐心等候...</p>"
	Response.Flush

	i=0
	n=Request.QueryString("n")
	n=Int(n)
	m=True

	For Each f1 in fc

		Set objXmlVerChk=New PluginSapper_CheckVersionViaXML

		If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "Plugin.xml") Then

			objXmlVerChk.XmlDataLocal=(LoadFromFile(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/plugin.xml","utf-8"))

			If i>=n Then
				Response.Write "<p class=""status-box"" id=""checking"">>>> 插件: " & objXmlVerChk.Item_Name_Local & " 正在检查中...</p>"
				Response.Flush
			Else
				If fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "verchk.xml") Then
					t="<b class=""notice"">发现新版本!</b>"
				ElseIf fso.FileExists(BlogPath & "/ZB_USERS/PLUGIN/" & f1.name & "/" & "error.log") Then
					t="<span class=""somehow"">该插件不支持在线更新.</span>"
				Else
					t="<span class=""allright"">未发现新版本.</span>"
				End If

				Response.Write "<p class=""status-box"">>>> 插件: " & objXmlVerChk.Item_Name_Local & " " & t & "</p>"
				Response.Flush
			End If

			If i>=n Then
				s=getHTTPPage(Resource_URL & f1.name & "/verchk.xml")

				objXmlVerChk.XmlDataWeb=(s)

				If objXmlVerChk.UpdateNeeded Then

					t="<b>" & objXmlVerChk.Item_Name_Local & "</b> 检查完毕, <b class=""notice"">发现新版本!</b>"

				ElseIf s=False Then

					t="<b>" & objXmlVerChk.Item_Name_Local & "</b> 检查完毕, <span class=""somehow"">该插件不支持在线更新.</span>"

				Else

					t="<b>" & objXmlVerChk.Item_Name_Local & "</b> 检查完毕, <span class=""allright"">未发现新版本.</span>"

				End If

				i=i+1

				Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('checking').style.display = 'none';}catch(e){};</script>"
				Response.Write "<p class=""download-box"">" & t & "</p>"
				Response.Write "<script>setTimeout(""self.location.href='Xml_ChkVer.asp?act=check&n=" & i & "'"",3000);</script>"
				Response.Flush

				Call SaveToFile(BlogPath & "/ZB_USERS/PLUGIN/PluginSapper/Export/log.txt",f1.name,"utf-8",False)

				m=False

				Set objXmlVerChk=Nothing
				Exit For

			End If

			i=i+1

		End If

		Set objXmlVerChk=Nothing

	Next

	If m Then

		Response.Write "<p class=""status-box"">!! 所有插件已检查完成!</p>"

		Response.Write "<form name=""edit"" method=""get"" action=""#"" class=""status-box"">"
		Response.Write "<p><input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /> <input onclick=""self.location.href='Xml_ChkVer.asp';"" type=""button"" class=""button"" value=""查看需要更新的插件"" title=""查看更新结果"" /></p>"
		Response.Write "</form>"

		Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading2').style.display = 'none';}catch(e){};</script>"
	End If

End If

Set fso = nothing
Err.Clear
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