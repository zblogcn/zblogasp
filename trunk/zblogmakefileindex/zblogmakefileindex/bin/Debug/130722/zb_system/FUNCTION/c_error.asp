<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_error.asp
'// 开始时间:    2004.07.25
'// 最后修改:    
'// 备    注:    错误显示页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="c_function.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<link rel="stylesheet" rev="stylesheet" href="../css/admin.css" type="text/css" media="screen" />
	<title><%=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG045%></title>
</head>
<body class="short">

<div class="bg">
<div id="wrapper">
  <div class="logo"><img src="../image/admin/none.gif" title="Z-Blog" alt="Z-Blog"/></div>
  <div class="login">
	<form id="frmLogin" method="post" action="">
	  <div class="divHeader"><%=ZC_MSG045%></div>

<%
	Dim a,b
	a=Request.QueryString("errorid")
	b=Request.QueryString("number")
	Call CheckParameter(a,"int",0)
	Call CheckParameter(b,"int",0)
	Response.Write "<p>" & ZC_MSG098 & ":" & ZVA_ErrorMsg(a) & "</p>"

	If b<>0 Then
		Response.Write "<p>" & ZC_MSG076 & ":" & "" & a & "</p>"
		Response.Write "<p>" & ZC_MSG016 & ":" & "<br/>" & TransferHTML(Request.QueryString("description"),"[html-format]") & "</p>"
		Response.Write "<p>" & TransferHTML(Request.QueryString("source"),"[html-format]") & "</p>"
	End If
		Response.Write "<p><br/></p>"
	If CheckRegExp(Request.QueryString("sourceurl"),"[homepage]")=True Then
		Response.Write "<p style='text-align:right;'><a href=""" & TransferHTML(Request.QueryString("sourceurl"),"[html-format]") & """>" & ZC_MSG207 & "</a></p>"
	Else
		Response.Write "<p style='text-align:right;'><a href=""" & GetCurrentHost() & """>" & ZC_MSG207 & "</a></p>"
	End If

	If a=6 Then
		Response.Write "<p style='text-align:right;'><a href=""../cmd.asp?act=login"" target=""_top"">"& ZC_MSG009 & "</a></p>"
	End If
%>

    </form>
  </div>
</div>
</div>
</body>
</html>
<%
If Err.Number<>0 Then
	Response.Redirect GetCurrentHost() & "zb_system/function/c_error.asp"
End If
%>