<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Response.CacheControl="no-cache" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%
Call System_Initialize()


If Request("wid") Then
	Dim viewid,ZC_PU_WEIXIN_TITLE,ZC_PU_WEIXIN_TIME,ZC_PU_WEIXIN_POSTER,ZC_PU_WEIXIN_CONTENT
	viewid = Request("wid")
	
	Dim objArticle
	Set objArticle=New TArticle
	
	If objArticle.LoadInfoByID(viewid) Then
		ZC_PU_WEIXIN_TIME = objArticle.PostTime
		ZC_PU_WEIXIN_TITLE = objArticle.TITLE
		'ZC_PU_WEIXIN_POSTER = Users(objArticle.AuthorID).Name
		ZC_PU_WEIXIN_CONTENT = objArticle.Content
		
	End If
Else
	Response.End()
End If



%>
<?xml version="1.0" encoding="UTF-8"?> <!DOCTYPE html PUBLIC "-//WAPFORUM//DTD XHTML Mobile 1.0//EN" "http://www.wapforum.org/DTD/xhtml-mobile10.dtd"> <html xmlns="http://www.w3.org/1999/xhtml"><head><title><%=ZC_PU_WEIXIN_TITLE%></title>  </head> <body class="read_url"> <div id="main" class="ru_container" style="position:relative;top:0;left:0;bottom:0;right:0;overflow:hidden;margin:20px auto;"> <div class="ru_article"> <h1 class="ru_title"><%=ZC_PU_WEIXIN_TITLE%></h1> <div class="ru_time"><%=ZC_PU_WEIXIN_TIME%> <a class="btn_mailme" target="_blank" href=""><%=ZC_PU_WEIXIN_POSTER%></a></div> <div id="ru_content" class="ru_content"><%=ZC_PU_WEIXIN_CONTENT%> </div> </div> </div> <div class="ru_footer" style="display:block;position:static;margin:0 auto;"> <p class="info_exception" style="display:block;color:#999;margin:0;"> 本文章内容为博客《<%=ZC_BLOG_TITLE%>》通过 <a href="" target="_blank">微信搜索插件</a>生成，以供文章在微信上查看获得最佳体验。 <a target="_blank" href="">免责声明</a> <br />Powered By <a target="_blank" href="http://www.rainbowsoft.org/">Z-Blog & 未寒</a></p> </div> <style type="text/css">body,p,h1,h6,img{margin:0;padding:0}img{border:0 none}a{color:#1e5494}body.read_url{background-color:#f2f2f2;*background-color:#fff;font-family:Verdana,Arial,Helvetica}.ru_container{display:table;*display:block;background-color:#fff;border:1px solid #aaa;*border-color:#fff;box-shadow:0 2px 2px rgba(0,0,0,0.1);width:820px;_width:820px}.ru_article{margin:30px 40px 40px;min-height:500px;_height:500px}.ru_title{font-size:24px;font-weight:bold;font-family:微软雅黑;line-height:1.2}.ru_time{font-size:12px;color:#999;margin:10px 0 20px}a.btn_mailme{margin-left:10px;text-decoration:underline;color:#999}.ru_content{line-height:2}.ru_att{position:relative;border:1px solid #ddd;border-radius:5px;margin:20px 0;padding:15px 0 15px 15px}.ru_att h6{position:absolute;top:-8px;left:21px;padding:0 3px;background-color:#fff;font-size:12px}a.att_panel{padding:5px;border-radius:5px;display:inline-block;cursor:pointer;text-decoration:none;width:224px}a.att_panel:hover{background-color:#f4f4f4}a.att_panel:active{background-color:#eee}.att_panel img{float:left}.att_info{margin-left:36px;margin-right:5px;color:#999;font-size:12px;_width:180px;_overflow:hidden}.att_name{color:#000;margin-bottom:3px;max-width:190px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.ru_footer{font-size:12px;width:800px;*width:auto;*background-color:#eee;*border-top:1px solid #bbb;*text-align:center}.info_exception{line-height:1.8;padding:0 0 20px;*padding-top:20px}.info_exception a{color:#999;text-decoration:underline}body.ru_mob{background-color:#fff}.ru_mob .ru_container{border-width:0;box-shadow:none;width:auto;margin:0}.ru_mob .ru_article{margin:10px}.ru_mob .ru_att{padding:10px 0 10px 10px}.ru_mob .ru_footer{width:auto;padding:20px 10px;background-color:#eee;border-top:1px solid #bbb;text-align:left}.ru_mob .info_exception{padding-bottom:0}.ru_mob .ru_footer a{white-space:nowrap}.relief_top .ru_footer{position:absolute;top:0;left:50%;width:770px;margin:20px 0 0 -385px}.relief_top .ru_container{margin-top:80px}.url404{font-size:14px;line-height:1.3;border:1px solid #bbb;background-color:#fff;position:absolute;left:50%;margin-left:-169px;margin-top:50px}.url404 .icon_info_b{float:left;display:block;margin:30px 0 0 30px}.url404 .url404_intro{margin:30px 30px 30px 75px}</style> </body> </html> 