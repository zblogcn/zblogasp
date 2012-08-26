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
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->

<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)
Call WindsPhoto_Initialize

Dim strWP_SCRIPT_TYPE
strWP_SCRIPT_TYPE = Replace(Replace(Request.Form("strWP_SCRIPT_TYPE"), VBCRLF, ""), VBLF, "")
WP_Config.Write "WP_SCRIPT_TYPE",strWP_SCRIPT_TYPE

Dim strWP_WATERMARK_TYPE
strWP_WATERMARK_TYPE = Replace(Replace(Request.Form("strWP_WATERMARK_TYPE"), VBCRLF, ""), VBLF, "")
WP_Config.Write "WP_WATERMARK_TYPE",strWP_WATERMARK_TYPE

Dim strWP_ORDER_BY
strWP_ORDER_BY = Replace(Replace(Request.Form("strWP_ORDER_BY"), VBCRLF, ""), VBLF, "")
WP_Config.Write "WP_ORDER_BY",strWP_ORDER_BY

Dim numWP_UPLOAD_FILESIZE
numWP_UPLOAD_FILESIZE = Request.Form("numWP_UPLOAD_FILESIZE")
Call SaveValueForSetting(strContent, TRUE, "Numeric", "WP_UPLOAD_FILESIZE", numWP_UPLOAD_FILESIZE)

Dim strWP_UPLOAD_DIR
strWP_UPLOAD_DIR = Request.Form("strWP_UPLOAD_DIR")
WP_Config.Write "WP_UPLOAD_DIR",strWP_UPLOAD_DIR

Dim strWP_UPLOAD_DIRBY
strWP_UPLOAD_DIRBY = Request.Form("strWP_UPLOAD_DIRBY")
WP_Config.Write "WP_UPLOAD_DIRBY",strWP_UPLOAD_DIRBY

Dim strWP_JPEG_FONTBOLD
strWP_JPEG_FONTBOLD = Request.Form("strWP_JPEG_FONTBOLD")
WP_Config.Write "WP_JPEG_FONTBOLD",strWP_JPEG_FONTBOLD

Dim strWP_JPEG_FONTQUALITY
strWP_JPEG_FONTQUALITY = Request.Form("strWP_JPEG_FONTQUALITY")
WP_Config.Write "WP_JPEG_FONTQUALITY",strWP_JPEG_FONTQUALITY

Dim strWP_JPEG_FONTSIZE
strWP_JPEG_FONTSIZE = Request.Form("strWP_JPEG_FONTSIZE")
WP_Config.Write "WP_JPEG_FONTSIZE",strWP_JPEG_FONTSIZE

Dim strWP_JPEG_FONTCOLOR
strWP_JPEG_FONTCOLOR = Request.Form("strWP_JPEG_FONTCOLOR")
WP_Config.Write "WP_JPEG_FONTCOLOR",strWP_JPEG_FONTCOLOR

Dim strWP_WATERMARK_TEXT
strWP_WATERMARK_TEXT = Request.Form("strWP_WATERMARK_TEXT")
WP_Config.Write "WP_WATERMARK_TEXT",strWP_WATERMARK_TEXT

Dim strWP_ALBUM_NAME
strWP_ALBUM_NAME = Request.Form("strWP_ALBUM_NAME")
WP_Config.Write "WP_ALBUM_NAME",strWP_ALBUM_NAME

Dim strWP_WATERMARK_WIDTH_POSITION
strWP_WATERMARK_WIDTH_POSITION = Request.Form("strWP_WATERMARK_WIDTH_POSITION")
WP_Config.Write "WP_WATERMARK_WIDTH_POSITION",strWP_WATERMARK_WIDTH_POSITION

Dim strWP_WATERMARK_HEIGHT_POSITION
strWP_WATERMARK_HEIGHT_POSITION = Request.Form("strWP_WATERMARK_HEIGHT_POSITION")
WP_Config.Write "WP_WATERMARK_HEIGHT_POSITION",strWP_WATERMARK_HEIGHT_POSITION

Dim numWP_SMALL_WIDTH
numWP_SMALL_WIDTH = Request.Form("numWP_SMALL_WIDTH")
Call SaveValueForSetting(strContent, TRUE, "Numeric", "WP_SMALL_WIDTH", numWP_SMALL_WIDTH)

Dim numWP_SMALL_HEIGHT
numWP_SMALL_HEIGHT = Request.Form("numWP_SMALL_HEIGHT")
Call SaveValueForSetting(strContent, TRUE, "Numeric", "WP_SMALL_HEIGHT", numWP_SMALL_HEIGHT)

Dim strWP_WATERMARK_LOGO
strWP_WATERMARK_LOGO = Request.Form("strWP_WATERMARK_LOGO")
WP_Config.Write "WP_WATERMARK_LOGO",strWP_WATERMARK_LOGO

Dim strWP_WATERMARK_ALPHA
strWP_WATERMARK_ALPHA = Request.Form("strWP_WATERMARK_ALPHA")
WP_Config.Write "WP_WATERMARK_ALPHA",strWP_WATERMARK_ALPHA

Dim numWP_LIST_HEIGHT
numWP_LIST_HEIGHT = Request.Form("numWP_LIST_HEIGHT")
Call SaveValueForSetting(strContent, TRUE, "Numeric", "WP_LIST_HEIGHT", numWP_LIST_HEIGHT)

Dim numWP_LIST_WIDTH
numWP_LIST_WIDTH = Request.Form("numWP_LIST_WIDTH")
Call SaveValueForSetting(strContent, TRUE, "Numeric", "WP_LIST_WIDTH", numWP_LIST_WIDTH)

Dim numWP_INDEX_PAGERCOUNT
numWP_INDEX_PAGERCOUNT = Request.Form("numWP_INDEX_PAGERCOUNT")
Call SaveValueForSetting(strContent, TRUE, "Numeric", "WP_INDEX_PAGERCOUNT", numWP_INDEX_PAGERCOUNT)

Dim numWP_SMALL_PAGERCOUNT
numWP_SMALL_PAGERCOUNT = Request.Form("numWP_SMALL_PAGERCOUNT")
Call SaveValueForSetting(strContent, TRUE, "Numeric", "WP_SMALL_PAGERCOUNT", numWP_SMALL_PAGERCOUNT)

Dim numWP_LIST_PAGERCOUNT
numWP_LIST_PAGERCOUNT = Request.Form("numWP_LIST_PAGERCOUNT")
Call SaveValueForSetting(strContent, TRUE, "Numeric", "WP_LIST_PAGERCOUNT", numWP_LIST_PAGERCOUNT)

Dim strWP_SUB_DOMAIN
strWP_SUB_DOMAIN = Request.Form("strWP_SUB_DOMAIN")
WP_Config.Write "WP_SUB_DOMAIN",strWP_SUB_DOMAIN

Dim strWP_ALBUM_INTRO
strWP_ALBUM_INTRO = Request.Form("strWP_ALBUM_INTRO")
WP_Config.Write "WP_ALBUM_INTRO",strWP_ALBUM_INTRO

Dim strWP_UPLOAD_RENAME
strWP_UPLOAD_RENAME = Request.Form("strWP_UPLOAD_RENAME")
WP_Config.Write "WP_UPLOAD_RENAME",strWP_UPLOAD_RENAME

Dim strWP_WATERMARK_AUTO
strWP_WATERMARK_AUTO = Request.Form("strWP_WATERMARK_AUTO")
WP_Config.Write "WP_WATERMARK_AUTO",strWP_WATERMARK_AUTO

Dim numWP_BLOGPHOTO_ID
numWP_BLOGPHOTO_ID = Request.Form("numWP_BLOGPHOTO_ID")
Call SaveValueForSetting(strContent, TRUE, "Numeric", "WP_BLOGPHOTO_ID", numWP_BLOGPHOTO_ID)

Dim strWP_HIDE_DIVFILESND
strWP_HIDE_DIVFILESND = Request.Form("strWP_HIDE_DIVFILESND")
WP_Config.Write "WP_HIDE_DIVFILESND",strWP_HIDE_DIVFILESND

Call WP_Config.Save

Call SetBlogHint_Custom("√ 设置成功.")

Response.Redirect "admin_setting.asp"

Call System_Terminate()

If Err.Number<>0 Then
    Call ShowError(0)
End If
%>