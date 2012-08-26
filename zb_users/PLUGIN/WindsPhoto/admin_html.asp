<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备   注:     WindsPhoto
'// 最后修改：   2010.6.10
'// 最后版本:    2.7.1
'///////////////////////////////////////////////////////////////////////////////
%>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../p_config.asp" -->
<!-- #include file="data/conn.asp" -->
<!-- #include file="function.asp" -->

<%
Call System_Initialize

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

LoadGlobeCache

Dim ArtList
Set ArtList = New TArticleList

ArtList.LoadCache

If LoadFromFile(BlogPath & "Themes/" & ZC_BLOG_THEME & "/Template/wp_index.html", "utf-8") = "" Then
    ArtList.template = "TAGS"
Else
    ArtList.template = "WP_INDEX"
End If

ArtList.Title = WP_ALBUM_NAME

ArtList.Build
    Dim Html, AddedHtml
    Html = ArtList.html
    AddedHtml = "<link rel=""alternate"" type=""application/rss+xml"" href="""& WP_SUB_DOMAIN &"rss.asp"" title=""订阅我的相册"" />" & VBCRLF
    AddedHtml = AddedHtml & "<script type=""text/javascript"" src="""& WP_SUB_DOMAIN &"script/windsphoto.js""></script>" & VBCRLF
    AddedHtml = AddedHtml & "<link rel=""stylesheet"" href="""& WP_SUB_DOMAIN &"images/windsphoto.css"" type=""text/css"" media=""screen"" />" & VBCRLF & "<title>"
    Html = Replace(Html, "<title>", AddedHtml)
    Html = Replace(Html, "<#BlogTitle#>", WP_ALBUM_NAME)
    Html = Replace(Html, ">TagCloud</h2>", ">" & WP_ALBUM_NAME & "</h2>")
    Html = Replace(Html, "<#CUSTOM_TAGS_TITLE#>", WP_ALBUM_NAME)
    Html = Replace(Html, ">Powered By", ">Powered By <a href='http://photo.wilf.cn/' target='_blank' title='WindsPhoto官方网站'>WindsPhoto</a> &")
    Html = Replace(Html, "<#CUSTOM_TAGS#>", GetPhotoIndex())
    Call ClearGlobeCache
    Call LoadGlobeCache
    Call SaveToFile(BlogPath & "photo.html", Html, "utf-8", TRUE)
    Call SaveSortList()
    Call SetBlogHint_Custom("√ 操作成功.")
    Response.Redirect"admin_main.asp"

Set ArtList = Nothing

Set Conn = Nothing
'If Err.Number<>0 then
'	Call ShowError(0)
'End If
%>