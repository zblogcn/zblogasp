<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog 彩虹网志个人版
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    edit_setting.asp
'// 开始时间:    2005.03.16
'// 最后修改:    
'// 备    注:    编辑设置页
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
For Each sAction_Plugin_Edit_Setting_Begin in Action_Plugin_Edit_Setting_Begin
	If Not IsEmpty(sAction_Plugin_Edit_Setting_Begin) Then Call Execute(sAction_Plugin_Edit_Setting_Begin)
Next

'检查非法链接
Call CheckReference("")

'检查权限
If Not CheckRights("SettingMng") Then Call ShowError(6)

GetCategory()
GetUser()

Dim EditArticle

BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG247

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
	<title><%=BlogTitle%></title>
</head>
<body>
			<div id="divMain">
<div class="Header"><%=ZC_MSG247%></div>
<%
	Response.Write "<div class=""SubMenu"">" & Response_Plugin_SettingMng_SubMenu & "</div>"
%>
<form method="post" action="../cmd.asp?act=SettingSav">
<div id="divMain2">
<% Call GetBlogHint() %>
<ul>
	<li class="tabs-selected"><a href="#fragment-1"><span><%=ZC_MSG105%></span></a></li>
	<li><a href="#fragment-2"><span><%=ZC_MSG173%></span></a></li>
	<li><a href="#fragment-3"><span><%=ZC_MSG186%></span></a></li>
	<li><a href="#fragment-4"><span><%=ZC_MSG281%></span></a></li>
	<li><a href="#fragment-5"><span><%=ZC_MSG195%></span></a></li>
	<li><a href="#fragment-6"><span><%=ZC_MSG215%></span></a></li>
</ul>
<%

	Function SplitNameAndNote(s)

		Dim i,j

		i=InStr(s,"(")
		j=InStr(s,")")

		If i>0 And j>0 Then 
			SplitNameAndNote="<p  align='left'>·" & Left(s,i-1) & ""
			SplitNameAndNote=SplitNameAndNote & "<p>" & Mid(s,i+1,Len(s)-i+1-2) & "</p></p>"
		Else
			SplitNameAndNote="<p  align='left'>·" & s & "</p>"
		End If
		
	End Function


	Dim i,j
	Dim tmpSng

	tmpSng=LoadFromFile(BlogPath & "zb_users/c_custom.asp","utf-8")

	Dim strZC_BLOG_HOST
	Dim strZC_BLOG_TITLE
	Dim strZC_BLOG_SUBTITLE
	Dim strZC_BLOG_NAME
	Dim strZC_BLOG_SUB_NAME
	Dim strZC_BLOG_CSS
	Dim strZC_BLOG_COPYRIGHT
	Dim strZC_BLOG_MASTER
	Dim strZC_BLOG_THEME

	Call LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_HOST",strZC_BLOG_HOST)
	Call LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_TITLE",strZC_BLOG_TITLE)
	Call LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_SUBTITLE",strZC_BLOG_SUBTITLE)
	Call LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_NAME",strZC_BLOG_NAME)
	Call LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_SUB_NAME",strZC_BLOG_SUB_NAME)
	Call LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_CSS",strZC_BLOG_CSS)
	Call LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_COPYRIGHT",strZC_BLOG_COPYRIGHT)
	Call LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_MASTER",strZC_BLOG_MASTER)
	Call LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_THEME",strZC_BLOG_THEME)


	strZC_BLOG_HOST=TransferHTML(strZC_BLOG_HOST,"[html-format]")
	strZC_BLOG_TITLE=TransferHTML(strZC_BLOG_TITLE,"[html-format]")
	strZC_BLOG_SUBTITLE=TransferHTML(strZC_BLOG_SUBTITLE,"[html-format]")
	strZC_BLOG_NAME=TransferHTML(strZC_BLOG_NAME,"[html-format]")
	strZC_BLOG_SUB_NAME=TransferHTML(strZC_BLOG_SUB_NAME,"[html-format]")
	strZC_BLOG_CSS=TransferHTML(strZC_BLOG_CSS,"[html-format]")
	strZC_BLOG_COPYRIGHT=TransferHTML(strZC_BLOG_COPYRIGHT,"[html-format]")
	strZC_BLOG_MASTER=TransferHTML(strZC_BLOG_MASTER,"[html-format]")
	strZC_BLOG_THEME=TransferHTML(strZC_BLOG_THEME,"[html-format]")

	Response.Write "<div class=""tabs-div"" style='border:none;padding:0px;margin:0;' id=""fragment-1"">"
	Response.Write "<table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>"
	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG104) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_HOST"" name=""edtZC_BLOG_HOST"" style=""width:95%"" type=""text"" value=""" & strZC_BLOG_HOST & """ /></p></td></tr>"
	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG091) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_NAME"" name=""edtZC_BLOG_NAME"" style=""width:95%"" type=""text"" value=""" & strZC_BLOG_NAME & """ /></p></td></tr>"
	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG092) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_SUB_NAME"" name=""edtZC_BLOG_SUB_NAME"" style=""width:95%""  type=""text"" value=""" & strZC_BLOG_SUB_NAME & """ /></p></td></tr>"
	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG093) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_TITLE"" name=""edtZC_BLOG_TITLE""style=""width:95%""  type=""text"" value=""" & strZC_BLOG_TITLE &""" /></p></td></tr>"
	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG094) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_SUBTITLE"" name=""edtZC_BLOG_SUBTITLE"" style=""width:95%""  type=""text"" value=""" & strZC_BLOG_SUBTITLE & """ /></p></td></tr>"
	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG293) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_THEME"" name=""edtZC_BLOG_THEME"" style=""width:95%"" type=""text"" value=""" & strZC_BLOG_THEME & """ /></p></td></tr>"
	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG095) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_CSS"" name=""edtZC_BLOG_CSS"" style=""width:95%"" type=""text"" value=""" & strZC_BLOG_CSS & """ /></p></td></tr>"
	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG096) & "</td><td style=""width:68%""><p><textarea rows=""4"" id=""edtZC_BLOG_COPYRIGHT"" name=""edtZC_BLOG_COPYRIGHT"" style=""width:95%"" type=""text"" >" & strZC_BLOG_COPYRIGHT & "</textarea></p></td></tr>"
	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG097) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_MASTER"" name=""edtZC_BLOG_MASTER"" style=""width:95%""  type=""text"" value=""" & strZC_BLOG_MASTER & """ /></p></td></tr>"
	Response.Write "</table>"
	Response.Write "</div>"



	Response.Write "<div class=""tabs-div"" style='border:none;padding:0px;margin:0;' id=""fragment-2"">"
	Response.Write "<table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>"
	tmpSng=LoadFromFile(BlogPath & "zb_users/c_option.asp","utf-8")


	Dim strZC_BLOG_CLSID
	If LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_CLSID",strZC_BLOG_CLSID) Then
		strZC_BLOG_CLSID=TransferHTML(strZC_BLOG_CLSID,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG174) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_CLSID"" name=""edtZC_BLOG_CLSID"" style=""width:95%"" type=""text"" value=""" & strZC_BLOG_CLSID & """ /></p></td></tr>"
	End If

	Dim strZC_TIME_ZONE
	If LoadValueForSetting(tmpSng,True,"String","ZC_TIME_ZONE",strZC_TIME_ZONE) Then
		strZC_TIME_ZONE=TransferHTML(strZC_TIME_ZONE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG175) & "</td><td style=""width:68%""><p><input id=""edtZC_TIME_ZONE"" name=""edtZC_TIME_ZONE"" style=""width:95%"" type=""text"" value=""" & strZC_TIME_ZONE & """ /></p></td></tr>"
	End If

	Dim strZC_HOST_TIME_ZONE
	If LoadValueForSetting(tmpSng,True,"String","ZC_HOST_TIME_ZONE",strZC_HOST_TIME_ZONE) Then
		strZC_HOST_TIME_ZONE=TransferHTML(strZC_HOST_TIME_ZONE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG303) & "</td><td style=""width:68%""><p><input id=""edtZC_HOST_TIME_ZONE"" name=""edtZC_HOST_TIME_ZONE"" style=""width:95%"" type=""text"" value=""" & strZC_HOST_TIME_ZONE & """ /></p></td></tr>"
	End If

	Dim strZC_BLOG_LANGUAGE
	If LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_LANGUAGE",strZC_BLOG_LANGUAGE) Then
		strZC_BLOG_LANGUAGE=TransferHTML(strZC_BLOG_LANGUAGE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG176) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_LANGUAGE"" name=""edtZC_BLOG_LANGUAGE"" style=""width:95%"" type=""text"" value=""" & strZC_BLOG_LANGUAGE & """ /></p></td></tr>"
	End If


	Dim strZC_UPDATE_INFO_URL
	If LoadValueForSetting(tmpSng,True,"String","ZC_UPDATE_INFO_URL",strZC_UPDATE_INFO_URL) Then
		strZC_UPDATE_INFO_URL=TransferHTML(strZC_UPDATE_INFO_URL,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG290) & "</td><td style=""width:68%""><p><input id=""edtZC_UPDATE_INFO_URL"" name=""edtZC_UPDATE_INFO_URL"" style=""width:95%"" type=""text"" value=""" & strZC_UPDATE_INFO_URL & """/></p></td></tr>"
	End If

	Dim strZC_BLOG_WEBEDIT
	If LoadValueForSetting(tmpSng,True,"String","ZC_BLOG_WEBEDIT",strZC_BLOG_WEBEDIT) Then
		strZC_BLOG_WEBEDIT=TransferHTML(strZC_BLOG_WEBEDIT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG180) & "</td><td style=""width:68%""><p><input id=""edtZC_BLOG_WEBEDIT"" name=""edtZC_BLOG_WEBEDIT"" style=""width:95%"" type=""text"" value=""" & strZC_BLOG_WEBEDIT & """ /></p></td></tr>"
	End If

	Dim strZC_UPLOAD_FILETYPE
	If LoadValueForSetting(tmpSng,True,"String","ZC_UPLOAD_FILETYPE",strZC_UPLOAD_FILETYPE) Then
		strZC_UPLOAD_FILETYPE=TransferHTML(strZC_UPLOAD_FILETYPE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG183) & "</td><td style=""width:68%""><p><input id=""edtZC_UPLOAD_FILETYPE"" name=""edtZC_UPLOAD_FILETYPE"" style=""width:95%"" type=""text"" value=""" & strZC_UPLOAD_FILETYPE & """ /></p></td></tr>"
	End If

	Dim strZC_UPLOAD_FILESIZE
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_UPLOAD_FILESIZE",strZC_UPLOAD_FILESIZE) Then
		strZC_UPLOAD_FILESIZE=TransferHTML(strZC_UPLOAD_FILESIZE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG184) & "</td><td style=""width:68%""><p><input id=""edtZC_UPLOAD_FILESIZE"" name=""edtZC_UPLOAD_FILESIZE"" style=""width:95%"" type=""text"" value=""" & strZC_UPLOAD_FILESIZE & """ /></p></td></tr>"
	End If

	Dim strZC_UPLOAD_DIRBYMONTH
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UPLOAD_DIRBYMONTH",strZC_UPLOAD_DIRBYMONTH) Then
		strZC_UPLOAD_DIRBYMONTH=TransferHTML(strZC_UPLOAD_DIRBYMONTH,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG292) & "</td><td style=""width:68%""><p><input id=""edtZC_UPLOAD_DIRBYMONTH"" name=""edtZC_UPLOAD_DIRBYMONTH"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UPLOAD_DIRBYMONTH),"checked","")&" value=""" & strZC_UPLOAD_DIRBYMONTH & """ class=""pointer"" ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_RSS_EXPORT_WHOLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_RSS_EXPORT_WHOLE",strZC_RSS_EXPORT_WHOLE) Then
		strZC_RSS_EXPORT_WHOLE=TransferHTML(strZC_RSS_EXPORT_WHOLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG226) & "</td><td style=""width:68%""><p><input id=""edtZC_RSS_EXPORT_WHOLE"" name=""edtZC_RSS_EXPORT_WHOLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_RSS_EXPORT_WHOLE),"checked","")&" value=""" & strZC_RSS_EXPORT_WHOLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_GUEST_REVERT_COMMENT_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_GUEST_REVERT_COMMENT_ENABLE",strZC_GUEST_REVERT_COMMENT_ENABLE) Then
		strZC_GUEST_REVERT_COMMENT_ENABLE=TransferHTML(strZC_GUEST_REVERT_COMMENT_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG283) & "</td><td style=""width:68%""><p><input id=""edtZC_GUEST_REVERT_COMMENT_ENABLE"" name=""edtZC_GUEST_REVERT_COMMENT_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_GUEST_REVERT_COMMENT_ENABLE),"checked","")&" value=""" & strZC_GUEST_REVERT_COMMENT_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_COMMENT_TURNOFF
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_COMMENT_TURNOFF",strZC_COMMENT_TURNOFF) Then
		strZC_COMMENT_TURNOFF=TransferHTML(strZC_COMMENT_TURNOFF,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG262) & "</td><td style=""width:68%""><p><input id=""edtZC_COMMENT_TURNOFF"" name=""edtZC_COMMENT_TURNOFF"" style="""" type=""checkbox"" "&IIf(CBool(strZC_COMMENT_TURNOFF),"checked","")&" value=""" & strZC_COMMENT_TURNOFF & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_TRACKBACK_TURNOFF
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_TRACKBACK_TURNOFF",strZC_TRACKBACK_TURNOFF) Then
		strZC_TRACKBACK_TURNOFF=TransferHTML(strZC_TRACKBACK_TURNOFF,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG263) & "</td><td style=""width:68%""><p><input id=""edtZC_TRACKBACK_TURNOFF"" name=""edtZC_TRACKBACK_TURNOFF"" style="""" type=""checkbox"" "&IIf(CBool(strZC_TRACKBACK_TURNOFF),"checked","")&" value=""" & strZC_TRACKBACK_TURNOFF & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If


	Dim strZC_GUESTBOOK_CONTENT
	If LoadValueForSetting(tmpSng,True,"String","ZC_GUESTBOOK_CONTENT",strZC_GUESTBOOK_CONTENT) Then
		strZC_GUESTBOOK_CONTENT=TransferHTML(strZC_GUESTBOOK_CONTENT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG276) & "</td><td style=""width:68%""><p><textarea rows=""6"" id=""edtZC_GUESTBOOK_CONTENT"" name=""edtZC_GUESTBOOK_CONTENT"" style=""width:95%"" type=""text"" >" & strZC_GUESTBOOK_CONTENT & "</textarea></p></td></tr>"
	End If

	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div class=""tabs-div"" style='border:none;padding:0px;margin:0;' id=""fragment-3"">"
	Response.Write "<table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>"

	Dim strZC_MSG_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_MSG_COUNT",strZC_MSG_COUNT) Then
		strZC_MSG_COUNT=TransferHTML(strZC_MSG_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG187) & "</td><td style=""width:68%""><p><input id=""edtZC_MSG_COUNT"" name=""edtZC_MSG_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_MSG_COUNT & """ /></p></td></tr>"
	End If

	Dim strZC_ARCHIVE_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_ARCHIVE_COUNT",strZC_ARCHIVE_COUNT) Then
		strZC_ARCHIVE_COUNT=TransferHTML(strZC_ARCHIVE_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG188) & "</td><td style=""width:68%""><p><input id=""edtZC_ARCHIVE_COUNT"" name=""edtZC_ARCHIVE_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_ARCHIVE_COUNT & """ /></p></td></tr>"
	End If

	Dim strZC_PREVIOUS_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_PREVIOUS_COUNT",strZC_PREVIOUS_COUNT) Then
		strZC_PREVIOUS_COUNT=TransferHTML(strZC_PREVIOUS_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG189) & "</td><td style=""width:68%""><p><input id=""edtZC_PREVIOUS_COUNT"" name=""edtZC_PREVIOUS_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_PREVIOUS_COUNT & """ /></p></td></tr>"
	End If

	Dim strZC_DISPLAY_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_DISPLAY_COUNT",strZC_DISPLAY_COUNT) Then
		strZC_DISPLAY_COUNT=TransferHTML(strZC_DISPLAY_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG190) & "</td><td style=""width:68%""><p><input id=""edtZC_DISPLAY_COUNT"" name=""edtZC_DISPLAY_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_DISPLAY_COUNT & """ /></p></td></tr>"
	End If

	Dim strZC_MANAGE_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_MANAGE_COUNT",strZC_MANAGE_COUNT) Then
		strZC_MANAGE_COUNT=TransferHTML(strZC_MANAGE_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG191) & "</td><td style=""width:68%""><p><input id=""edtZC_MANAGE_COUNT"" name=""edtZC_MANAGE_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_MANAGE_COUNT & """ /></p></td></tr>"
	End If

	Dim strZC_RSS2_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_RSS2_COUNT",strZC_RSS2_COUNT) Then
		strZC_RSS2_COUNT=TransferHTML(strZC_RSS2_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG192) & "</td><td style=""width:68%""><p><input id=""edtZC_RSS2_COUNT"" name=""edtZC_RSS2_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_RSS2_COUNT & """ /></p></td></tr>"
	End If

	Dim strZC_SEARCH_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_SEARCH_COUNT",strZC_SEARCH_COUNT) Then
		strZC_SEARCH_COUNT=TransferHTML(strZC_SEARCH_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG193) & "</td><td style=""width:68%""><p><input id=""edtZC_SEARCH_COUNT"" name=""edtZC_SEARCH_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_SEARCH_COUNT & """ /></p></td></tr>"
	End If

	Dim strZC_PAGEBAR_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_PAGEBAR_COUNT",strZC_PAGEBAR_COUNT) Then
		strZC_PAGEBAR_COUNT=TransferHTML(strZC_PAGEBAR_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG194) & "</td><td style=""width:68%""><p><input id=""edtZC_PAGEBAR_COUNT"" name=""edtZC_PAGEBAR_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_PAGEBAR_COUNT & """ /></p></td></tr>"
	End If

	Dim strZC_USE_NAVIGATE_ARTICLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_USE_NAVIGATE_ARTICLE",strZC_USE_NAVIGATE_ARTICLE) Then
		strZC_USE_NAVIGATE_ARTICLE=TransferHTML(strZC_USE_NAVIGATE_ARTICLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG209) & "</td><td style=""width:68%""><p><input id=""edtZC_USE_NAVIGATE_ARTICLE"" name=""edtZC_USE_NAVIGATE_ARTICLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_USE_NAVIGATE_ARTICLE),"checked","")&" value=""" & strZC_USE_NAVIGATE_ARTICLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_MUTUALITY_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_MUTUALITY_COUNT",strZC_MUTUALITY_COUNT) Then
		strZC_MUTUALITY_COUNT=TransferHTML(strZC_MUTUALITY_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG230) & "</td><td style=""width:68%""><p><input id=""edtZC_MUTUALITY_COUNT"" name=""edtZC_MUTUALITY_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_MUTUALITY_COUNT & """ /></p></td></tr>"
	End If

	Dim strZC_COMMENT_REVERSE_ORDER_EXPORT
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_COMMENT_REVERSE_ORDER_EXPORT",strZC_COMMENT_REVERSE_ORDER_EXPORT) Then
		strZC_COMMENT_REVERSE_ORDER_EXPORT=TransferHTML(strZC_COMMENT_REVERSE_ORDER_EXPORT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG271) & "</td><td style=""width:68%""><p><input id=""edtZC_COMMENT_REVERSE_ORDER_EXPORT"" name=""edtZC_COMMENT_REVERSE_ORDER_EXPORT"" style="""" type=""checkbox"" "&IIf(CBool(strZC_COMMENT_REVERSE_ORDER_EXPORT),"checked","")&" value=""" & strZC_COMMENT_REVERSE_ORDER_EXPORT & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_COMMENT_VERIFY_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_COMMENT_VERIFY_ENABLE",strZC_COMMENT_VERIFY_ENABLE) Then
		strZC_COMMENT_VERIFY_ENABLE=TransferHTML(strZC_COMMENT_VERIFY_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG185) & "</td><td style=""width:68%""><p><input id=""edtZC_COMMENT_VERIFY_ENABLE"" name=""edtZC_COMMENT_VERIFY_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_COMMENT_VERIFY_ENABLE),"checked","")&" value=""" & strZC_COMMENT_VERIFY_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_VERIFYCODE_STRING
	If LoadValueForSetting(tmpSng,True,"String","ZC_VERIFYCODE_STRING",strZC_VERIFYCODE_STRING) Then
		strZC_VERIFYCODE_STRING=TransferHTML(strZC_VERIFYCODE_STRING,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG300) & "</td><td style=""width:68%""><p><input id=""edtZC_VERIFYCODE_STRING"" name=""edtZC_VERIFYCODE_STRING"" style=""width:95%"" type=""text"" value=""" & strZC_VERIFYCODE_STRING & """ /></p></td></tr>"
	End If

	Dim strZC_VERIFYCODE_WIDTH
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_VERIFYCODE_WIDTH",strZC_VERIFYCODE_WIDTH) Then
		strZC_VERIFYCODE_WIDTH=TransferHTML(strZC_VERIFYCODE_WIDTH,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG301) & "</td><td style=""width:68%""><p><input id=""edtZC_VERIFYCODE_WIDTH"" name=""edtZC_VERIFYCODE_WIDTH"" style=""width:95%"" type=""text"" value=""" & strZC_VERIFYCODE_WIDTH & """/></p></td></tr>"
	End If

	Dim strZC_VERIFYCODE_HEIGHT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_VERIFYCODE_HEIGHT",strZC_VERIFYCODE_HEIGHT) Then
		strZC_VERIFYCODE_HEIGHT=TransferHTML(strZC_VERIFYCODE_HEIGHT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG302) & "</td><td style=""width:68%""><p><input id=""edtZC_VERIFYCODE_HEIGHT"" name=""edtZC_VERIFYCODE_HEIGHT"" style=""width:95%"" type=""text"" value=""" & strZC_VERIFYCODE_HEIGHT & """/></p></td></tr>"
	End If

	Dim strZC_IMAGE_WIDTH
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_IMAGE_WIDTH",strZC_IMAGE_WIDTH) Then
		strZC_IMAGE_WIDTH=TransferHTML(strZC_IMAGE_WIDTH,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG171) & "</td><td style=""width:68%""><p><input id=""edtZC_IMAGE_WIDTH"" name=""edtZC_IMAGE_WIDTH"" style=""width:95%"" type=""text"" value=""" & strZC_IMAGE_WIDTH & """/></p></td></tr>"
	End If

	Dim strZC_RECENT_COMMENT_WORD_MAX
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_RECENT_COMMENT_WORD_MAX",strZC_RECENT_COMMENT_WORD_MAX) Then
		strZC_RECENT_COMMENT_WORD_MAX=TransferHTML(strZC_RECENT_COMMENT_WORD_MAX,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG206) & "</td><td style=""width:68%""><p><input id=""edtZC_RECENT_COMMENT_WORD_MAX"" name=""edtZC_RECENT_COMMENT_WORD_MAX"" style=""width:95%"" type=""text"" value=""" & strZC_RECENT_COMMENT_WORD_MAX & """/></p></td></tr>"
	End If

	Dim strZC_TAGS_DISPLAY_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_TAGS_DISPLAY_COUNT",strZC_TAGS_DISPLAY_COUNT) Then
		strZC_TAGS_DISPLAY_COUNT=TransferHTML(strZC_TAGS_DISPLAY_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(GetSettingFormNameWithDefault("ZC_MSG319","Tags Count in SliderBar")) & "</td><td style=""width:68%""><p><input id=""edtZC_TAGS_DISPLAY_COUNT"" name=""edtZC_TAGS_DISPLAY_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_TAGS_DISPLAY_COUNT & """/></p></td></tr>"
	End If

	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div class=""tabs-div"" style='border:none;padding:0px;margin:0;' id=""fragment-4"">"
	Response.Write "<table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>"

	Dim strZC_STATIC_TYPE
	If LoadValueForSetting(tmpSng,True,"String","ZC_STATIC_TYPE",strZC_STATIC_TYPE) Then
		strZC_STATIC_TYPE=TransferHTML(strZC_STATIC_TYPE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG177) & "</td><td style=""width:68%""><p><input id=""edtZC_STATIC_TYPE"" name=""edtZC_STATIC_TYPE"" style=""width:95%"" type=""text"" value=""" & strZC_STATIC_TYPE & """ /></p></td></tr>"
	End If

	Dim strZC_STATIC_DIRECTORY
	If LoadValueForSetting(tmpSng,True,"String","ZC_STATIC_DIRECTORY",strZC_STATIC_DIRECTORY) Then
		strZC_STATIC_DIRECTORY=TransferHTML(strZC_STATIC_DIRECTORY,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG178) & "</td><td style=""width:68%""><p><input id=""edtZC_STATIC_DIRECTORY"" name=""edtZC_STATIC_DIRECTORY"" style=""width:95%"" type=""text"" value=""" & strZC_STATIC_DIRECTORY & """ /></p></td></tr>"
	End If

	Dim strZC_CUSTOM_DIRECTORY_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_CUSTOM_DIRECTORY_ENABLE",strZC_CUSTOM_DIRECTORY_ENABLE) Then
		strZC_CUSTOM_DIRECTORY_ENABLE=TransferHTML(strZC_CUSTOM_DIRECTORY_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG278) & "<p><a href='http://wiki.rainbowsoft.org/doku.php?id=wiki:config:url' target='_blank'><font color='green'>Z-Wiki:配置Z-Blog的静态URL</font></a></p></td><td style=""width:68%""><p><input id=""edtZC_CUSTOM_DIRECTORY_ENABLE"" name=""edtZC_CUSTOM_DIRECTORY_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_CUSTOM_DIRECTORY_ENABLE),"checked","")&" value=""" & strZC_CUSTOM_DIRECTORY_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_CUSTOM_DIRECTORY_REGEX
	If LoadValueForSetting(tmpSng,True,"String","ZC_CUSTOM_DIRECTORY_REGEX",strZC_CUSTOM_DIRECTORY_REGEX) Then
		strZC_CUSTOM_DIRECTORY_REGEX=TransferHTML(strZC_CUSTOM_DIRECTORY_REGEX,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG279) & "</td><td style=""width:68%""><p><input id=""edtZC_CUSTOM_DIRECTORY_REGEX"" name=""edtZC_CUSTOM_DIRECTORY_REGEX"" style=""width:95%"" type=""text"" value=""" & strZC_CUSTOM_DIRECTORY_REGEX & """ /></p></td></tr>"
	End If

	Dim strZC_CUSTOM_DIRECTORY_ANONYMOUS
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_CUSTOM_DIRECTORY_ANONYMOUS",strZC_CUSTOM_DIRECTORY_ANONYMOUS) Then
		strZC_CUSTOM_DIRECTORY_ANONYMOUS=TransferHTML(strZC_CUSTOM_DIRECTORY_ANONYMOUS,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG280) & "</td><td style=""width:68%""><p><input id=""edtZC_CUSTOM_DIRECTORY_ANONYMOUS"" name=""edtZC_CUSTOM_DIRECTORY_ANONYMOUS"" style="""" type=""checkbox"" "&IIf(CBool(strZC_CUSTOM_DIRECTORY_ANONYMOUS),"checked","")&" value=""" & strZC_CUSTOM_DIRECTORY_ANONYMOUS & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_REBUILD_FILE_COUNT
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_REBUILD_FILE_COUNT",strZC_REBUILD_FILE_COUNT) Then
		strZC_REBUILD_FILE_COUNT=TransferHTML(strZC_REBUILD_FILE_COUNT,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG181) & "</td><td style=""width:68%""><p><input id=""edtZC_REBUILD_FILE_COUNT"" name=""edtZC_REBUILD_FILE_COUNT"" style=""width:95%"" type=""text"" value=""" & strZC_REBUILD_FILE_COUNT & """ /></p></td></tr>"
	End If

	Dim strZC_REBUILD_FILE_INTERVAL
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_REBUILD_FILE_INTERVAL",strZC_REBUILD_FILE_INTERVAL) Then
		strZC_REBUILD_FILE_INTERVAL=TransferHTML(strZC_REBUILD_FILE_INTERVAL,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG182) & "</td><td style=""width:68%""><p><input id=""edtZC_REBUILD_FILE_INTERVAL"" name=""edtZC_REBUILD_FILE_INTERVAL"" style=""width:95%"" type=""text"" value=""" & strZC_REBUILD_FILE_INTERVAL & """ /></p></td></tr>"
	End If


	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div class=""tabs-div"" style='border:none;padding:0px;margin:0;' id=""fragment-5"">"
	Response.Write "<table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>"

	Dim strZC_UBB_LINK_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UBB_LINK_ENABLE",strZC_UBB_LINK_ENABLE) Then
		strZC_UBB_LINK_ENABLE=TransferHTML(strZC_UBB_LINK_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG196) & "</td><td style=""width:68%""><p><input id=""edtZC_UBB_LINK_ENABLE"" name=""edtZC_UBB_LINK_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UBB_LINK_ENABLE),"checked","")&" value=""" & strZC_UBB_LINK_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_UBB_FONT_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UBB_FONT_ENABLE",strZC_UBB_FONT_ENABLE) Then
		strZC_UBB_FONT_ENABLE=TransferHTML(strZC_UBB_FONT_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG197) & "</td><td style=""width:68%""><p><input id=""edtZC_UBB_FONT_ENABLE"" name=""edtZC_UBB_FONT_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UBB_FONT_ENABLE),"checked","")&" value=""" & strZC_UBB_FONT_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_UBB_CODE_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UBB_CODE_ENABLE",strZC_UBB_CODE_ENABLE) Then
		strZC_UBB_CODE_ENABLE=TransferHTML(strZC_UBB_CODE_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG198) & "</td><td style=""width:68%""><p><input id=""edtZC_UBB_CODE_ENABLE"" name=""edtZC_UBB_CODE_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UBB_CODE_ENABLE),"checked","")&" value=""" & strZC_UBB_CODE_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_UBB_FACE_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UBB_FACE_ENABLE",strZC_UBB_FACE_ENABLE) Then
		strZC_UBB_FACE_ENABLE=TransferHTML(strZC_UBB_FACE_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG199) & "</td><td style=""width:68%""><p><input id=""edtZC_UBB_FACE_ENABLE"" name=""edtZC_UBB_FACE_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UBB_FACE_ENABLE),"checked","")&" value=""" & strZC_UBB_FACE_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_UBB_IMAGE_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UBB_IMAGE_ENABLE",strZC_UBB_IMAGE_ENABLE) Then
		strZC_UBB_IMAGE_ENABLE=TransferHTML(strZC_UBB_IMAGE_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG200) & "</td><td style=""width:68%""><p><input id=""edtZC_UBB_IMAGE_ENABLE"" name=""edtZC_UBB_IMAGE_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UBB_IMAGE_ENABLE),"checked","")&" value=""" & strZC_UBB_IMAGE_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_UBB_MEDIA_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UBB_MEDIA_ENABLE",strZC_UBB_MEDIA_ENABLE) Then
		strZC_UBB_MEDIA_ENABLE=TransferHTML(strZC_UBB_MEDIA_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG201) & "</td><td style=""width:68%""><p><input id=""edtZC_UBB_MEDIA_ENABLE"" name=""edtZC_UBB_MEDIA_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UBB_MEDIA_ENABLE),"checked","")&" value=""" & strZC_UBB_MEDIA_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_UBB_FLASH_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UBB_FLASH_ENABLE",strZC_UBB_FLASH_ENABLE) Then
		strZC_UBB_FLASH_ENABLE=TransferHTML(strZC_UBB_FLASH_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG202) & "</td><td style=""width:68%""><p><input id=""edtZC_UBB_FLASH_ENABLE"" name=""edtZC_UBB_FLASH_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UBB_FLASH_ENABLE),"checked","")&" value=""" & strZC_UBB_FLASH_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_UBB_TYPESET_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UBB_TYPESET_ENABLE",strZC_UBB_TYPESET_ENABLE) Then
		strZC_UBB_TYPESET_ENABLE=TransferHTML(strZC_UBB_TYPESET_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG203) & "</td><td style=""width:68%""><p><input id=""edtZC_UBB_TYPESET_ENABLE"" name=""edtZC_UBB_TYPESET_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UBB_TYPESET_ENABLE),"checked","")&" value=""" & strZC_UBB_TYPESET_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_UBB_AUTOLINK_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UBB_AUTOLINK_ENABLE",strZC_UBB_AUTOLINK_ENABLE) Then
		strZC_UBB_AUTOLINK_ENABLE=TransferHTML(strZC_UBB_AUTOLINK_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG204) & "</td><td style=""width:68%""><p><input id=""edtZC_UBB_AUTOLINK_ENABLE"" name=""edtZC_UBB_AUTOLINK_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UBB_AUTOLINK_ENABLE),"checked","")&" value=""" & strZC_UBB_AUTOLINK_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	'Dim strZC_UBB_AUTOKEY_ENABLE
	'If LoadValueForSetting(tmpSng,True,"Boolean","ZC_UBB_AUTOKEY_ENABLE",strZC_UBB_AUTOKEY_ENABLE) Then
	'	strZC_UBB_AUTOKEY_ENABLE=TransferHTML(strZC_UBB_AUTOKEY_ENABLE,"[html-format]")
	'	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG205) & "</td><td style=""width:68%""><p><input id=""edtZC_UBB_AUTOKEY_ENABLE"" name=""edtZC_UBB_AUTOKEY_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_UBB_AUTOKEY_ENABLE),"checked","")&" value=""" & strZC_UBB_AUTOKEY_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	'End If

	Dim strZC_COMMENT_NOFOLLOW_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_COMMENT_NOFOLLOW_ENABLE",strZC_COMMENT_NOFOLLOW_ENABLE) Then
		strZC_COMMENT_NOFOLLOW_ENABLE=TransferHTML(strZC_COMMENT_NOFOLLOW_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG207) & "</td><td style=""width:68%""><p><input id=""edtZC_COMMENT_NOFOLLOW_ENABLE"" name=""edtZC_COMMENT_NOFOLLOW_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_COMMENT_NOFOLLOW_ENABLE),"checked","")&" value=""" & strZC_COMMENT_NOFOLLOW_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_JAPAN_TO_HTML
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_JAPAN_TO_HTML",strZC_JAPAN_TO_HTML) Then
		strZC_JAPAN_TO_HTML=TransferHTML(strZC_JAPAN_TO_HTML,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG208) & "</td><td style=""width:68%""><p><input id=""edtZC_JAPAN_TO_HTML"" name=""edtZC_JAPAN_TO_HTML"" style="""" type=""checkbox"" "&IIf(CBool(strZC_JAPAN_TO_HTML),"checked","")&" value=""" & strZC_JAPAN_TO_HTML & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If

	Dim strZC_EMOTICONS_FILENAME
	If LoadValueForSetting(tmpSng,True,"String","ZC_EMOTICONS_FILENAME",strZC_EMOTICONS_FILENAME) Then
		strZC_EMOTICONS_FILENAME=TransferHTML(strZC_EMOTICONS_FILENAME,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG235) & "</td><td style=""width:68%""><p><input id=""edtZC_EMOTICONS_FILENAME"" name=""edtZC_EMOTICONS_FILENAME"" style=""width:95%"" type=""text"" value=""" & strZC_EMOTICONS_FILENAME & """/></p></td></tr>"
	End If

	Dim strZC_EMOTICONS_FILESIZE
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_EMOTICONS_FILESIZE",strZC_EMOTICONS_FILESIZE) Then
		strZC_EMOTICONS_FILESIZE=TransferHTML(strZC_EMOTICONS_FILESIZE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG234) & "</td><td style=""width:68%""><p><input id=""edtZC_EMOTICONS_FILESIZE"" name=""edtZC_EMOTICONS_FILESIZE"" style=""width:95%"" type=""text"" value=""" & strZC_EMOTICONS_FILESIZE & """ /></p></td></tr>"
	End If

	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "<div class=""tabs-div"" style='border:none;padding:0px;margin:0;' id=""fragment-6"">"
	Response.Write "<table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>"

	'Dim strZC_IE_DISPLAY_WAP
	'If LoadValueForSetting(tmpSng,True,"Boolean","ZC_IE_DISPLAY_WAP",strZC_IE_DISPLAY_WAP) Then
	'	strZC_IE_DISPLAY_WAP=TransferHTML(strZC_IE_DISPLAY_WAP,"[html-format]")
	'	Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG216) & "</td><td style=""width:68%""><p><input id=""edtZC_IE_DISPLAY_WAP"" name=""edtZC_IE_DISPLAY_WAP"" style="""" type=""checkbox"" "&IIf(CBool(strZC_IE_DISPLAY_WAP),"checked","")&" value=""" & strZC_IE_DISPLAY_WAP & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	'End If

	Dim strZC_DISPLAY_COUNT_WAP
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_DISPLAY_COUNT_WAP",strZC_DISPLAY_COUNT_WAP) Then
		strZC_DISPLAY_COUNT_WAP=TransferHTML(strZC_DISPLAY_COUNT_WAP,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG217) & "</td><td style=""width:68%""><p><input id=""edtZC_DISPLAY_COUNT_WAP"" name=""edtZC_DISPLAY_COUNT_WAP"" style=""width:95%"" type=""text"" value=""" & strZC_DISPLAY_COUNT_WAP & """ /></p></td></tr>"
	End If

	Dim strZC_COMMENT_COUNT_WAP
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_COMMENT_COUNT_WAP",strZC_COMMENT_COUNT_WAP) Then
		strZC_COMMENT_COUNT_WAP=TransferHTML(strZC_COMMENT_COUNT_WAP,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG218) & "</td><td style=""width:68%""><p><input id=""edtZC_COMMENT_COUNT_WAP"" name=""edtZC_COMMENT_COUNT_WAP"" style=""width:95%"" type=""text"" value=""" & strZC_COMMENT_COUNT_WAP & """ /></p></td></tr>"
	End If

	Dim strZC_PAGEBAR_COUNT_WAP
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_PAGEBAR_COUNT_WAP",strZC_PAGEBAR_COUNT_WAP) Then
		strZC_PAGEBAR_COUNT_WAP=TransferHTML(strZC_PAGEBAR_COUNT_WAP,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG219) & "</td><td style=""width:68%""><p><input id=""edtZC_PAGEBAR_COUNT_WAP"" name=""edtZC_PAGEBAR_COUNT_WAP"" style=""width:95%"" type=""text"" value=""" & strZC_PAGEBAR_COUNT_WAP & """ /></p></td></tr>"
	End If

	Dim strZC_SINGLE_SIZE_WAP
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_SINGLE_SIZE_WAP",strZC_SINGLE_SIZE_WAP) Then
		strZC_SINGLE_SIZE_WAP=TransferHTML(strZC_SINGLE_SIZE_WAP,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG220) & "</td><td style=""width:68%""><p><input id=""edtZC_SINGLE_SIZE_WAP"" name=""edtZC_SINGLE_SIZE_WAP"" style=""width:95%"" type=""text"" value=""" & strZC_SINGLE_SIZE_WAP & """ /></p></td></tr>"
	End If

	Dim strZC_SINGLE_PAGEBAR_COUNT_WAP
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_SINGLE_PAGEBAR_COUNT_WAP",strZC_SINGLE_PAGEBAR_COUNT_WAP) Then
		strZC_SINGLE_PAGEBAR_COUNT_WAP=TransferHTML(strZC_SINGLE_PAGEBAR_COUNT_WAP,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG221) & "</td><td style=""width:68%""><p><input id=""edtZC_SINGLE_PAGEBAR_COUNT_WAP"" name=""edtZC_SINGLE_PAGEBAR_COUNT_WAP"" style=""width:95%"" type=""text"" value=""" & strZC_SINGLE_PAGEBAR_COUNT_WAP & """ /></p></td></tr>"
	End If

	Dim strZC_COMMENT_PAGEBAR_COUNT_WAP
	If LoadValueForSetting(tmpSng,True,"Numeric","ZC_COMMENT_PAGEBAR_COUNT_WAP",strZC_COMMENT_PAGEBAR_COUNT_WAP) Then
		strZC_COMMENT_PAGEBAR_COUNT_WAP=TransferHTML(strZC_COMMENT_PAGEBAR_COUNT_WAP,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG222) & "</td><td style=""width:68%""><p><input id=""edtZC_COMMENT_PAGEBAR_COUNT_WAP"" name=""edtZC_COMMENT_PAGEBAR_COUNT_WAP"" style=""width:95%"" type=""text"" value=""" & strZC_COMMENT_PAGEBAR_COUNT_WAP & """ /></p></td></tr>"
	End If

	Dim strZC_FILENAME_WAP
	If LoadValueForSetting(tmpSng,True,"String","ZC_FILENAME_WAP",strZC_FILENAME_WAP) Then
		strZC_FILENAME_WAP=TransferHTML(strZC_FILENAME_WAP,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG223) & "</td><td style=""width:68%""><p><input id=""edtZC_FILENAME_WAP"" name=""edtZC_FILENAME_WAP"" style=""width:95%"" type=""text"" value=""" & strZC_FILENAME_WAP & """/></p></td></tr>"
	End If

	Dim strZC_WAPCOMMENT_ENABLE
	If LoadValueForSetting(tmpSng,True,"Boolean","ZC_WAPCOMMENT_ENABLE",strZC_WAPCOMMENT_ENABLE) Then
		strZC_WAPCOMMENT_ENABLE=TransferHTML(strZC_WAPCOMMENT_ENABLE,"[html-format]")
		Response.Write "<tr><td style='width:32%'>" & SplitNameAndNote(ZC_MSG304) & "</td><td style=""width:68%""><p><input id=""edtZC_WAPCOMMENT_ENABLE"" name=""edtZC_WAPCOMMENT_ENABLE"" style="""" type=""checkbox"" "&IIf(CBool(strZC_WAPCOMMENT_ENABLE),"checked","")&" value=""" & strZC_WAPCOMMENT_ENABLE & """ ONCLICK=""ChangeValue(this);""/></p></td></tr>"
	End If 

	Response.Write "</table>"
	Response.Write "</div>"

	Response.Write "<p><br/><input type=""submit"" class=""button"" value="""& ZC_MSG087 &""" id=""btnPost"" onclick='' /></p>"

%>


			</div></form></div>
<script language="javascript">


$(document).ready(function(){
	$("#divMain2").tabs({ fxFade: true, fxSpeed: 'fast' });
	$("input[type=text],textarea").width($("body").width()*0.55);

	//斑马线
	var tables=document.getElementsByTagName("table");
	var b=false;
	for (var j = 0; j < tables.length; j++){

		var cells = tables[j].getElementsByTagName("tr");

		//cells[0].className="color3";
		b=false;
		for (var i = 0; i < cells.length; i++){
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

});



function ChangeValue(obj){

	if (obj.value=="True")
	{
	obj.value="False";
	return true;
	}

	if (obj.value=="False")
	{
	obj.value="True";
	return true;
	}
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