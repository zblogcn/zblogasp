<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:  Z-Blog 1.8 其它版本未知
'// 插件制作:  狼的旋律(http://www.wilf.cn) 
'// 备    注:  WindsPhoto
'// 最后修改： 2011.8.22
'// 最后版本:  2.7.3
'///////////////////////////////////////////////////////////////////////////////
%>
<%' Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%
Call System_Initialize()

Call CheckReference("")

If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

'检查ASPJPEG组件,并把是否存在aspjpeg缓存到设置中,有用
Set Jpeg = Server.CreateObject("Persits.Jpeg")
If -2147221005 = Err Or Jpeg.Expires<Now() Then
    Call SetBlogHint_Custom("!! 当前服务器没有ASPJPEG组件,你只可以在无法生产缩略图和水印功能的情况下继续使用本相册.</a>")
    If WP_IF_ASPJPEG = "1" Then
        Dim strContent
        strContent = LoadFromFile(BlogPath & "/plugin/WindsPhoto/include.asp", "utf-8")
        Call SaveValueForSetting(strContent, TRUE, "String", "WP_IF_ASPJPEG", "0")
        Call SaveToFile(BlogPath & "/plugin/WindsPhoto/include.asp", strContent, "utf-8", FALSE)
        Response.Redirect "admin_setting.asp"
    End If
ElseIf WP_IF_ASPJPEG = "0" Then
    Dim strContent1
    strContent1 = LoadFromFile(BlogPath & "/plugin/WindsPhoto/include.asp", "utf-8")
    Call SaveValueForSetting(strContent1, TRUE, "String", "WP_IF_ASPJPEG", "1")
    Call SaveToFile(BlogPath & "/plugin/WindsPhoto/include.asp", strContent1, "utf-8", FALSE)
    Response.Redirect "admin_setting.asp"
End If
Set Jpeg = Nothing

If WP_SUB_DOMAIN = "" Then
    Dim strContent3
    strContent3 = LoadFromFile(BlogPath & "/plugin/WindsPhoto/include.asp", "utf-8")
    Call SaveValueForSetting(strContent3, TRUE, "String", "WP_SUB_DOMAIN", ZC_BLOG_HOST&"plugin/windsphoto/")
    Call SaveToFile(BlogPath & "/plugin/WindsPhoto/include.asp", strContent3, "utf-8", FALSE)
    Response.Redirect "admin_setting.asp"
End If

BlogTitle = "WindsPhoto 相册设置"
%>

<%

Dim tmpSng

tmpSng = LoadFromFile(BlogPath & "/plugin/WindsPhoto/include.asp", "utf-8")

Dim strWP_SCRIPT_TYPE
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_SCRIPT_TYPE", strWP_SCRIPT_TYPE)

Dim strWP_WATERMARK_TYPE
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_WATERMARK_TYPE", strWP_WATERMARK_TYPE)

Dim strWP_ORDER_BY
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_ORDER_BY", strWP_ORDER_BY)

Dim numWP_UPLOAD_FILESIZE
Call LoadValueForSetting(tmpSng, TRUE, "Numeric", "WP_UPLOAD_FILESIZE", numWP_UPLOAD_FILESIZE)

Dim strWP_UPLOAD_DIR
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_UPLOAD_DIR", strWP_UPLOAD_DIR)

Dim strWP_UPLOAD_DIRBY
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_UPLOAD_DIRBY", strWP_UPLOAD_DIRBY)

Dim strWP_JPEG_FONTQUALITY
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_JPEG_FONTQUALITY", strWP_JPEG_FONTQUALITY)

Dim strWP_JPEG_FONTBOLD
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_JPEG_FONTBOLD", strWP_JPEG_FONTBOLD)

Dim strWP_JPEG_FONTSIZE
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_JPEG_FONTSIZE", strWP_JPEG_FONTSIZE)

Dim strWP_JPEG_FONTCOLOR
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_JPEG_FONTCOLOR", strWP_JPEG_FONTCOLOR)

Dim strWP_WATERMARK_TEXT
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_WATERMARK_TEXT", strWP_WATERMARK_TEXT)

Dim strWP_ALBUM_NAME
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_ALBUM_NAME", strWP_ALBUM_NAME)

Dim strWP_WATERMARK_WIDTH_POSITION
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_WATERMARK_WIDTH_POSITION", strWP_WATERMARK_WIDTH_POSITION)

Dim strWP_WATERMARK_HEIGHT_POSITION
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_WATERMARK_HEIGHT_POSITION", strWP_WATERMARK_HEIGHT_POSITION)

Dim numWP_SMALL_WIDTH
Call LoadValueForSetting(tmpSng, TRUE, "Numeric", "WP_SMALL_WIDTH", numWP_SMALL_WIDTH)

Dim numWP_SMALL_HEIGHT
Call LoadValueForSetting(tmpSng, TRUE, "Numeric", "WP_SMALL_HEIGHT", numWP_SMALL_HEIGHT)

Dim numWP_LIST_WIDTH
Call LoadValueForSetting(tmpSng, TRUE, "Numeric", "WP_LIST_WIDTH", numWP_LIST_WIDTH)

Dim numWP_LIST_HEIGHT
Call LoadValueForSetting(tmpSng, TRUE, "Numeric", "WP_LIST_HEIGHT", numWP_LIST_HEIGHT)

Dim strWP_WATERMARK_LOGO
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_WATERMARK_LOGO", strWP_WATERMARK_LOGO)

Dim strWP_WATERMARK_ALPHA
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_WATERMARK_ALPHA", strWP_WATERMARK_ALPHA)

Dim numWP_SMALL_PAGERCOUNT
Call LoadValueForSetting(tmpSng, TRUE, "Numeric", "WP_SMALL_PAGERCOUNT", numWP_SMALL_PAGERCOUNT)

Dim numWP_LIST_PAGERCOUNT
Call LoadValueForSetting(tmpSng, TRUE, "Numeric", "WP_LIST_PAGERCOUNT", numWP_LIST_PAGERCOUNT)

Dim strWP_SUB_DOMAIN
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_SUB_DOMAIN", strWP_SUB_DOMAIN)

Dim strWP_ALBUM_INTRO
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_ALBUM_INTRO", strWP_ALBUM_INTRO)

Dim strWP_UPLOAD_RENAME
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_UPLOAD_RENAME", strWP_UPLOAD_RENAME)

Dim strWP_WATERMARK_AUTO
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_WATERMARK_AUTO", strWP_WATERMARK_AUTO)

Dim numWP_INDEX_PAGERCOUNT
Call LoadValueForSetting(tmpSng, TRUE, "Numeric", "WP_INDEX_PAGERCOUNT", numWP_INDEX_PAGERCOUNT)

Dim numWP_BLOGPHOTO_ID
Call LoadValueForSetting(tmpSng, TRUE, "Numeric", "WP_BLOGPHOTO_ID", numWP_BLOGPHOTO_ID)

Dim strWP_HIDE_DIVFILESND
Call LoadValueForSetting(tmpSng, TRUE, "String", "WP_HIDE_DIVFILESND", strWP_HIDE_DIVFILESND)
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta name="robots" content="noindex,nofollow"/>
	<link rel="stylesheet" rev="stylesheet" href="../../CSS/admin.css" type="text/css" media="screen" />
	<script language="JavaScript" src="../../script/common.js" type="text/javascript"></script>
	<script language="JavaScript" src="../../script/jquery.tabs.pack.js" type="text/javascript"></script>
	<link rel="stylesheet" href="../../CSS/jquery.tabs.css" type="text/css" media="print, projection, screen">
	<!--[if lte IE 7]>
	<link rel="stylesheet" href="../../CSS/jquery.tabs-ie.css" type="text/css" media="projection, screen">
	<![endif]-->
	<link rel="stylesheet" href="../../CSS/jquery.bettertip.css" type="text/css" media="screen">
	<script language="JavaScript" src="../../script/jquery.bettertip.pack.js" type="text/javascript"></script>
	<title><%=BlogTitle%></title>
</head>
<body>
<div id="divMain">
	<div class="Header">WindsPhoto 系统设置</div>
		<div class="SubMenu">
			<span class="m-left"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/admin_main.asp">相册管理</a></span>
			<span class="m-left"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/admin_addtype.asp">新建相册</a></span>
			<span class="m-left m-now"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/admin_setting.asp">系统设置</a></span>	
			<span class="m-right"><a href="<%=ZC_BLOG_HOST%>cmd.asp?act=pluginMng">退出</a></span>
			<span class="m-right"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/help.asp">帮助说明</a></span>
			<span class="m-right"><a href="<%=ZC_BLOG_HOST%>PLUGIN/windsphoto/help.asp#more">更多功能</a></span>
		</div>
<form name="edit" method="post" action="admin_savesetting.asp">
<div id="divMain2">
<%Call GetBlogHint()%>
<ul>
	<li class="tabs-selected"><a href="#fragment-1"><span>参数设置</span></a></li>
	<li><a href="#fragment-2"><span>水印设置</span></a></li>
</ul>

<div class="tabs-div" style='border:none;padding:0px;margin:0;' id="fragment-1">
<table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>
<tr><td style='width:32%'><p align='left'>·相册名称</p></td><td style="width:68%"><p><input name="strWP_ALBUM_NAME" style="width:95%" type="text" value="<%=strWP_ALBUM_NAME%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·相册域名</p></td><td style="width:68%"><p><input name="strWP_SUB_DOMAIN" style="width:95%" type="text" value="<%=strWP_SUB_DOMAIN%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·贴图相册id(设置为0则不启用)new</p></td><td style="width:68%"><p><input name="numWP_BLOGPHOTO_ID" style="width:95%" type="text" value="<%=numWP_BLOGPHOTO_ID%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·隐藏Blog上传(上面不为0才起效)new</p></td><td style="width:68%"><p>	<input type="radio" name="strWP_HIDE_DIVFILESND" value="1" <%if WP_HIDE_DIVFILESND=1 then%>checked<%end if%> />是 <input type="radio" name="strWP_HIDE_DIVFILESND" value="0" <%if WP_HIDE_DIVFILESND=0 then%>checked<%end if%> />否</p></td></tr>

<tr><td style='width:32%'><p align='left'>·图片展示特效</p></td><td style="width:68%"><p>	<input type="radio" name="strWP_SCRIPT_TYPE" value="1" <%if WP_SCRIPT_TYPE=1 then%>checked<%end if%> />HighSlide <input type="radio" name="strWP_SCRIPT_TYPE" value="2" <%if WP_SCRIPT_TYPE=2 then%>checked<%end if%> />GreyBox <input type="radio" name="strWP_SCRIPT_TYPE" value="3" <%if WP_SCRIPT_TYPE=3 then%>checked<%end if%> />Lightbox <input type="radio" name="strWP_SCRIPT_TYPE" value="4" <%if WP_SCRIPT_TYPE=4 then%>checked<%end if%> />Thickbox</p></td></tr>

<tr><td style='width:32%'><p align='left'>·列表排序方式</p></td><td style="width:68%"><p>	<input type="radio" name="strWP_ORDER_BY" value="0" <%if WP_ORDER_BY=0 then%>checked<%end if%> />正序 <input type="radio" name="strWP_ORDER_BY" value="1" <%if WP_ORDER_BY=1 then%>checked<%end if%> />倒序</p></td></tr>

<tr><td style='width:32%'><p align='left'>·目录保存方式</p></td><td style="width:68%"><p>	<input type="radio" name="strWP_UPLOAD_DIRBY" value="1" <%if WP_UPLOAD_DIRBY=1 then%>checked<%end if%> />年/月	<input type="radio" name="strWP_UPLOAD_DIRBY" value="2" <%if WP_UPLOAD_DIRBY=2 then%>checked<%end if%> />分类id	<input type="radio" name="strWP_UPLOAD_DIRBY" value="0" <%if WP_UPLOAD_DIRBY=0 then%>checked<%end if%> />根目录</p></td></tr>

<tr><td style='width:32%'><p align='left'>·默认开启上传重命名 </p></td><td style="width:68%"><p>	<input type="radio" name="strWP_UPLOAD_RENAME" value="1" <%if WP_UPLOAD_RENAME=1 then%>checked<%end if%> />是 <input type="radio" name="strWP_UPLOAD_RENAME" value="0" <%if WP_UPLOAD_RENAME=0 then%>checked<%end if%> />否</p></td></tr>

<tr><td style='width:32%'><p align='left'>·上传文件限制</p></td><td style="width:68%"><p><input name="numWP_UPLOAD_FILESIZE" style="width:95%" type="text" value="<%=numWP_UPLOAD_FILESIZE%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·上传目录(结尾不加/) </p></td><td style="width:68%"><p><input name="strWP_UPLOAD_DIR" style="width:95%" type="text" value="<%=strWP_UPLOAD_DIR%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·首页显示相册数量 </p></td><td style="width:68%"><p><input name="numWP_INDEX_PAGERCOUNT" style="width:95%" type="text" value="<%=numWP_INDEX_PAGERCOUNT%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·缩略图模式分页</p></td><td style="width:68%"><p><input name="numWP_SMALL_PAGERCOUNT" style="width:95%" type="text" value="<%=numWP_SMALL_PAGERCOUNT%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·列表模式分页</p></td><td style="width:68%"><p><input name="numWP_LIST_PAGERCOUNT" style="width:95%" type="text" value="<%=numWP_LIST_PAGERCOUNT%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·列表模式图片限制</p></td><td style="width:68%"><p>高 <input name="numWP_LIST_WIDTH" style="width:42%" type="text" value="<%=numWP_LIST_WIDTH%>" /> 宽 <input name="numWP_LIST_HEIGHT" style="width:42%" type="text" value="<%=numWP_LIST_HEIGHT%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·缩略图转换大小</p></td><td style="width:68%"><p>高 <input name="numWP_SMALL_HEIGHT" style="width:42%" type="text" value="<%=numWP_SMALL_HEIGHT%>" /> 宽 <input name="numWP_SMALL_WIDTH" style="width:42%" type="text" value="<%=numWP_SMALL_WIDTH%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·相册文字介绍说明</p><p>支持HTML代码,可用&lt;br/&gt;'标签换行</p></td><td style="width:68%"><p><textarea name="strWP_ALBUM_INTRO" style="width:95%" rows="4" type="text" /><%=strWP_ALBUM_INTRO%></textarea></p></td></tr>

</table>
</div>

<div class="tabs-div" style='border:none;padding:0px;margin:0;' id="fragment-2">
<table width='100%' style='padding:0px;margin:1px;' cellspacing='0' cellpadding='0'>

<tr><td style='width:32%'><p align='left'>·默认开启水印</p></td><td style="width:68%"><p>	<input type="radio" name="strWP_WATERMARK_AUTO" value="1" <%if WP_WATERMARK_AUTO=1 then%>checked<%end if%> />是 <input type="radio" name="strWP_WATERMARK_AUTO" value="0" <%if WP_WATERMARK_AUTO=0 then%>checked<%end if%> />否</p></td></tr>

<tr><td style='width:32%'><p align='left'>·水印效果</p></td><td style="width:68%"><p><input type="radio" name="strWP_WATERMARK_TYPE" value="1" <%if WP_WATERMARK_TYPE=1 then%>checked<%end if%> />水印文字	<input type="radio" name="strWP_WATERMARK_TYPE" value="2" <%if WP_WATERMARK_TYPE=2 then%>checked<%end if%> />水印图片</p></td></tr>

<tr><td style='width:32%'><p align='left'>·水印水平位置</p></td><td style="width:68%"><p><input type="radio" name="strWP_WATERMARK_WIDTH_POSITION" value="left" <%if WP_WATERMARK_WIDTH_POSITION="left" then%>checked<%end if%> />左	<input type="radio" name="strWP_WATERMARK_WIDTH_POSITION" value="center" <%if WP_WATERMARK_WIDTH_POSITION="center" then%>checked<%end if%> />中	<input type="radio" name="strWP_WATERMARK_WIDTH_POSITION" value="right" <%if WP_WATERMARK_WIDTH_POSITION="right" then%>checked<%end if%> />右</p></td></tr>

<tr><td style='width:32%'><p align='left'>·水印垂直位置</p></td><td style="width:68%"><p><input type="radio" name="strWP_WATERMARK_HEIGHT_POSITION" value="top" <%if WP_WATERMARK_HEIGHT_POSITION="top" then%>checked<%end if%> />上	<input type="radio" name="strWP_WATERMARK_HEIGHT_POSITION" value="center" <%if WP_WATERMARK_HEIGHT_POSITION="center" then%>checked<%end if%> />中	<input type="radio" name="strWP_WATERMARK_HEIGHT_POSITION" value="bottom" <%if WP_WATERMARK_HEIGHT_POSITION="bottom" then%>checked<%end if%> />下</p></td></tr>

<tr><td style='width:32%'><p align='left'>·水印图片</p></td><td style="width:68%"><p><input name="strWP_WATERMARK_LOGO" style="width:95%" type="text" value="<%=strWP_WATERMARK_LOGO%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·水印透明</p></td><td style="width:68%"><p><input name="strWP_WATERMARK_ALPHA" style="width:95%" type="text" value="<%=strWP_WATERMARK_ALPHA%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·水印文字</p></td><td style="width:68%"><p><input name="strWP_WATERMARK_TEXT" style="width:95%" type="text" value="<%=strWP_WATERMARK_TEXT%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·输出质量</p></td><td style="width:68%"><p><input name="strWP_JPEG_FONTQUALITY" style="width:95%" type="text" value="<%=strWP_JPEG_FONTQUALITY%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·文字大小</p></td><td style="width:68%"><p><input name="strWP_JPEG_FONTSIZE" style="width:95%" type="text" value="<%=strWP_JPEG_FONTSIZE%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·文字颜色</p></td><td style="width:68%"><p><input name="strWP_JPEG_FONTCOLOR" style="width:95%" type="text" value="<%=strWP_JPEG_FONTCOLOR%>" /></p></td></tr>

<tr><td style='width:32%'><p align='left'>·是否粗体</p></td><td style="width:68%"><p>	<input type="radio" name="strWP_JPEG_FONTBOLD" value="true" <%if WP_JPEG_FONTBOLD="true" then%>checked<%end if%> />是	<input type="radio" name="strWP_JPEG_FONTBOLD" value="false" <%if WP_JPEG_FONTBOLD="false" then%>checked<%end if%> />否</p></td></tr>
</table>
</div>
<p><input type="submit" class="button" value="提交" id="btnPost" onclick='' /> <input type="reset" class="button" value="重置" id="btnPost" /></p>
</div>
</form>

<br><br><p align=center>Plugin Powered by <a href="http://www.wilf.cn" target="_blank">Wilf.cn</a></p>

</div>
<script language="javascript">
$(document).ready(function(){
	$("#divMain2").tabs({ fxFade: true, fxSpeed: 'fast' });
	//$("input[@type=text],textarea").width($("body").width()*0.55);

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

If Err.Number<>0 Then
  Call ShowError(0)
End If
%>