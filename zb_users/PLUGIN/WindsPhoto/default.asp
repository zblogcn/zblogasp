<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 spirit 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备    注:    WindsPhoto
'// 最后修改：   2010.6.10
'// 最后版本:    2.7.1
'///////////////////////////////////////////////////////////////////////////////
%>
<%' Option Explicit %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->

<!-- #include file="function.asp" -->

<%Call System_Initialize%><!-- #include file="data/conn.asp" --><%

LoadGlobeCache

Dim objArticle
Set objArticle = New TArticle
'objArticle.
objArticle.FType=ZC_POST_TYPE_PAGE
If GetTemplate("TEMPLATE_WP_INDEX")<>empty Then
    objArticle.template = "WP_INDEX"
End If



objArticle.Title = WP_ALBUM_NAME
objArticle.Content=GetPhotoIndex()
If objArticle.Export(ZC_DISPLAY_MODE_SYSTEMPAGE) Then
	objArticle.Build
    Dim Html, AddedHtml
    Html = objArticle.html
    AddedHtml = "<link rel=""alternate"" type=""application/rss+xml"" href="""& WP_SUB_DOMAIN &"rss.asp"" title=""订阅我的相册"" />" & VBCRLF
    AddedHtml = AddedHtml & "<script type=""text/javascript"" src="""& WP_SUB_DOMAIN &"script/windsphoto.js""></script>" & VBCRLF
    AddedHtml = AddedHtml & "<link rel=""stylesheet"" href="""& WP_SUB_DOMAIN &"images/windsphoto.css"" type=""text/css"" media=""screen"" />" & VBCRLF & "<title>"
    Html = Replace(Html, "<title>", AddedHtml)
    Html = Replace(Html, ">Powered By", ">Powered By <a href='http://photo.wilf.cn/' target='_blank' title='WindsPhoto官方网站'>WindsPhoto</a> &")
    Call ClearGlobeCache
    Call LoadGlobeCache
    Response.Write Html
End If

Set objArticle = Nothing

Set Conn = Nothing
'If Err.Number<>0 then
'	Call ShowError(0)
'End If
%>