<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// Z-Blog
'// 最后修改:    2011-8-3
'// 备    注:    WAP模块
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../c_option.asp" -->
<!-- #include file="../c_option_wap.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_event.asp" -->
<!-- #include file="../function/c_system_wap.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../plugin/p_config.asp" -->
<%
 ShowError_Custom="Call ShowError_WAP(id)"
%><?xml version="1.0" encoding="UTF-8"?> 
<!DOCTYPE html PUBLIC "-//WAPFORUM//DTD XHTML Mobile 1.0//EN" "http://www.wapforum.org/DTD/xhtml-mobile10.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link rel="stylesheet" href="<%=ZC_BLOG_HOST%>wap/style/wap.css" type="text/css" media="screen" charset="utf-8" />
<%
Call System_Initialize()
PubLic intPageCount
	Select Case Request.QueryString("act")
		Case "View"
			Call WapView()
		Case "Com"
			Call WapCom()
		Case "Main"
			Call WapMain()
		Case "Search"
			Call WapSearch()
		Case "Login"
			Call WapLogin()
		Case "Err"
			Call WapError()
		Case "Cate"
			Call WapCate()
		Case "Stat"
			Call WapStat()
		Case "Prev"
			Call WapPrev()			
		Case "AddCom"		
			Call WapAddCom(0)
		Case "PostCom"
			Call WapPostCom()
		Case "DelCom"
			Call WapDelCom()
		Case "AddArt"
		    Call WapEdtArt()
		Case "PostArt"
		    Call WapPostArt()
		Case "DelArt"
			Call WapDelArt()
		Case "Logout"
			Call WapLogout()
		Case Else
			Call WapMain()			
	End Select

Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>
<div id="ft">Powered By <a href="http://bbs.rainbowsoft.org">Z-Blog</a></div>
</body>
</html>