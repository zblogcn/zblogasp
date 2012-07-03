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
<!-- #include file="../../function/c_function_md5.asp" -->
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

Dim PageUrl,PageContent
Action=Request.QueryString("act")
PageUrl=Request.QueryString("url")
If PageUrl="" Then PageUrl=DownLoad_URL

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta name="robots" content="noindex,nofollow"/>
	<link rel="stylesheet" rev="stylesheet" href="../../CSS/admin.css" type="text/css" media="screen" />
	<link rel="stylesheet" rev="stylesheet" href="images/style.css" type="text/css" media="screen" />
	<title><%=BlogTitle%></title>
<%
	'为已安装的主题指定样式
	Response.Write "<style type=""text/css"">"& vbCrlf
	Dim fso, f, f1, fc, s
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "/THEMES/")
	Set fc = f.SubFolders
	For Each f1 in fc
		Response.Write "#theme"& MD5(LCase(f1.name)) &" {background:#F1FFFF url(""images/Installed.gif"");}"& vbCrlf
	Next
	Response.Write "</style>"
%>
</head>
<body>
<div id="divMain">
	<div class="Header">Theme Sapper - 获取更多主题 - 从服务器选择安装主题. <a href="help.asp#installonline" title="在线安装指南">[页面帮助]</a></div>
	<%Call SapperMenu("1")%>
<div id="divMain2">
<%
If Action <> "install" Then
	Call GetBlogHint()
	Response.Write "<p class=""hint hint_Teal""><font color=""Teal"">提示: 下面列出了""菠萝的海""里提供的主题资源, 您可以通过点击<b> [安装主题] </b>将您需要的主题安装到您的博客上.</font></p>"
End If
Response.Write "<div>"
Response.Write "<p id=""loading"">正在载入服务器数据, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
Response.Flush


PageContent=getHTTPPage(PageUrl)
PageContent=Replace(PageContent,"<#ZC_BLOG_HOST#>",ZC_BLOG_HOST)

Response.Write PageContent

Response.Write "<hr style=""clear:both;"" /><p><form name=""edit"" method=""get"" action=""#"" class=""status-box"">"
Response.Write "<p><input onclick=""self.location.href='ThemeList.asp';"" type=""button"" class=""button"" value=""返回主题管理"" title=""返回主题管理页"" /> <input onclick=""window.scrollTo(0,0);"" type=""button"" class=""button"" value=""TOP"" title=""返回页面顶部"" /></p>"
Response.Write "</form></p>"

Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('loading').style.display = 'none';</script>"
'*********************************************************
' 目的：    取得目标网页的html代码
'*********************************************************
function getHTTPPage(url)
dim Http,ServerConn
On Error Resume Next
dim j
For j=0 to 2
	set Http=server.createobject("Msxml2.ServerXMLHTTP")
	Http.setTimeouts SiteResolve*1000,SiteConnect*1000,SiteSend*1000,SiteReceive*1000
	Http.open "GET",url,false
	Http.send()

	if Http.readystate=4 then
		ServerConn = true
	else
		ServerConn = false
		set http=nothing
	end if

	if ServerConn then
		exit for
	end if
next
if err.number<>0 then err.Clear
if ServerConn = false then
	getHTTPPage = "<font color='red'> × 无法连接服务器!</font>"
	set http=nothing
	exit function
end if
getHTTPPage=Http.responseText
if http.Status=404 then
	getHTTPPage = "<font color='red'> × 服务器404错误!</font>"
end if
set http=nothing
end function
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