<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    
'// 开始时间:    
'// 最后修改:    
'// 备    注:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

'plugin node
For Each sAction_Plugin_Edit_Link_Begin in Action_Plugin_Edit_Link_Begin
	If Not IsEmpty(sAction_Plugin_Edit_Link_Begin) Then Call Execute(sAction_Plugin_Edit_Link_Begin)
Next

'检查非法链接
Call CheckReference("")

'检查权限
If Not CheckRights("LinkMng") Then Call ShowError(6)

GetCategory()
GetUser()

Dim EditArticle

BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG298

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<link rel="stylesheet" rev="stylesheet" href="../CSS/admin.css" type="text/css" media="screen" />
	<script language="JavaScript" src="../script/common.js" type="text/javascript"></script>
	<script language="JavaScript" src="../script/jquery.tabs.pack.js" type="text/javascript"></script>
	<link rel="stylesheet" href="../CSS/jquery.tabs.css" type="text/css" media="print, projection, screen">
	<!--[if lte IE 7]>
	<link rel="stylesheet" href="../CSS/jquery.tabs-ie.css" type="text/css" media="projection, screen">
	<![endif]-->
	<link rel="stylesheet" href="../CSS/jquery.bettertip.css" type="text/css" media="screen">
	<script language="JavaScript" src="../script/jquery.bettertip.pack.js" type="text/javascript"></script>
	<script language="JavaScript" src="../script/jquery.textarearesizer.compressed.js" type="text/javascript"></script>
	<title><%=BlogTitle%></title>
</head>
<body>
			<div id="divMain">
<div class="Header"><%=ZC_MSG298%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_LinkMng_SubMenu & "</div>"
%>
<form method="post" action="../cmd.asp?act=LinkSav">
<div id="divMain2">
<% Call GetBlogHint() %>
<ul>
	<li class="tabs-selected"><a href="#fragment-1"><span><%=ZC_MSG233%></span></a></li>
	<li><a href="#fragment-2"><span><%=ZC_MSG031%></span></a></li>
	<li><a href="#fragment-3"><span><%=ZC_MSG030%></span></a></li>
	<li><a href="#fragment-4"><span><%=ZC_MSG039%></span></a></li>
</ul>
<%
	Dim tpath

	Response.Write "<div class=""tabs-div"" id=""fragment-1"">"

	tpath="./ZB_USERS/INCLUDE/navbar.asp"

	Response.Write "<p>" & ZC_MSG170 & ": </p><p><INPUT TYPE=""text"" Value="""&unEscape(tpath)&""" style=""width:100%"" readonly></p>"
	Response.Write "<p></p>"
	Response.Write "<p><textarea class=""resizable""  style=""height:300px;width:100%"" name=""txaContent_Navbar"" id=""txaContent_Navbar"">"&TransferHTML(LoadFromFile(BlogPath & unEscape(tpath),"utf-8"),"[textarea]")&"</textarea></p>" & vbCrlf


	Response.Write "</div>"

	Response.Write "<div class=""tabs-div"" id=""fragment-2"">"

	tpath="./ZB_USERS/INCLUDE/link.asp"

	Response.Write "<p>" & ZC_MSG170 & ": </p><p><INPUT TYPE=""text"" Value="""&unEscape(tpath)&""" style=""width:100%"" readonly></p>"
	Response.Write "<p></p>"
	Response.Write "<p><textarea class=""resizable""   style=""height:300px;width:100%"" name=""txaContent_Link"" id=""txaContent_Link"">"&TransferHTML(LoadFromFile(BlogPath & unEscape(tpath),"utf-8"),"[textarea]")&"</textarea></p>" & vbCrlf

	Response.Write "</div>"
	Response.Write "<div class=""tabs-div"" id=""fragment-3"">"

	tpath="./ZB_USERS/INCLUDE/favorite.asp"

	Response.Write "<p>" & ZC_MSG170 & ": </p><p><INPUT TYPE=""text"" Value="""&unEscape(tpath)&""" style=""width:100%"" readonly></p>"
	Response.Write "<p></p>"
	Response.Write "<p><textarea class=""resizable""   style=""height:300px;width:100%"" name=""txaContent_Favorite"" id=""txaContent_Favorite"">"&TransferHTML(LoadFromFile(BlogPath & unEscape(tpath),"utf-8"),"[textarea]")&"</textarea></p>" & vbCrlf


	Response.Write "</div>"
	Response.Write "<div class=""tabs-div"" id=""fragment-4"">"

	tpath="./ZB_USERS/INCLUDE/misc.asp"

	Response.Write "<p>" & ZC_MSG170 & ": </p><p><INPUT TYPE=""text"" Value="""&unEscape(tpath)&""" style=""width:100%"" readonly></p>"
	Response.Write "<p></p>"
	Response.Write "<p><textarea class=""resizable""   style=""height:300px;width:100%"" name=""txaContent_Misc"" id=""txaContent_Misc"">"&TransferHTML(LoadFromFile(BlogPath & unEscape(tpath),"utf-8"),"[textarea]")&"</textarea></p>" & vbCrlf


	Response.Write "</div>"

	Response.Write "<p><br/><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='' /></p>"


%>


			</div></form></div>
<script language="javascript">


$(document).ready(function(){
	$("#divMain2").tabs({ fxFade: true, fxSpeed: 'fast' });

	/* jQuery textarea resizer plugin usage */
	$(document).ready(function() {
		$('textarea.resizable:not(.processed)').TextAreaResizer();
		$('iframe.resizable:not(.processed)').TextAreaResizer();
	});

});

</script>

</body>
</html>
<% 
Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>