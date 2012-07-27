<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    (zx.asd)&(sipo)&(月上之木)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    wap.asp
'// 开始时间:    2006-3-19
'// 最后修改:    2011-7-23
'// 备    注:    WAP模块
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Response.CacheControl="no-cache" %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_wap.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%
 ShowError_Custom="Call ShowError_WAP(id)"
%><?xml version="1.0" encoding="UTF-8"?> 
<!DOCTYPE html PUBLIC "-//WAPFORUM//DTD XHTML Mobile 1.0//EN" "http://www.wapforum.org/DTD/xhtml-mobile10.dtd"> 
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link rel="stylesheet" href="<%=ZC_BLOG_HOST%>zb_system/wap/style/wap.css" type="text/css" media="screen" charset="utf-8" />
<%
Call System_Initialize()

Call GetCategory()
Call GetUser()

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