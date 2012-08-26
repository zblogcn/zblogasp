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
Call WindsPhoto_Initialize
'检查ASPJPEG组件,并把是否存在aspjpeg缓存到设置中,有用
Set Jpeg = Server.CreateObject("Persits.Jpeg")
If -2147221005 = Err Or Jpeg.Expires<Now() Then
    Call SetBlogHint_Custom("!! 当前服务器没有ASPJPEG组件,你只可以在无法生产缩略图和水印功能的情况下继续使用本相册.</a>")
    If WP_IF_ASPJPEG = "1" Then
		WP_Config.Write "WP_IF_ASPJPEG", "0"
		WP_Config.Save
        Response.Redirect "admin_setting.asp"
    End If
ElseIf WP_IF_ASPJPEG = "0" Then
		WP_Config.Write "WP_IF_ASPJPEG", "1"
		WP_Config.Save
		   Response.Redirect "admin_setting.asp"
End If
Set Jpeg = Nothing

If WP_SUB_DOMAIN = "" Then
    WP_Config.Write "WP_SUB_DOMAIN", ZC_BLOG_HOST&"zb_users/plugin/windsphoto/"
	WP_Config.Save
    Response.Redirect "admin_setting.asp"
End If

BlogTitle = "WindsPhoto 相册设置"
%>

<%

Dim tmpSng

tmpSng = LoadFromFile(BlogPath & "/plugin/WindsPhoto/include.asp", "utf-8")

Dim strWP_SCRIPT_TYPE
strWP_SCRIPT_TYPE = WP_SCRIPT_TYPE
Dim strWP_WATERMARK_TYPE
strWP_WATERMARK_TYPE = WP_WATERMARK_TYPE
Dim strWP_ORDER_BY
strWP_ORDER_BY = WP_ORDER_BY
Dim numWP_UPLOAD_FILESIZE
numWP_UPLOAD_FILESIZE = WP_UPLOAD_FILESIZE
Dim strWP_UPLOAD_DIR
strWP_UPLOAD_DIR = WP_UPLOAD_DIR
Dim strWP_UPLOAD_DIRBY
strWP_UPLOAD_DIRBY = WP_UPLOAD_DIRBY
Dim strWP_JPEG_FONTQUALITY
strWP_JPEG_FONTQUALITY = WP_JPEG_FONTQUALITY
Dim strWP_JPEG_FONTBOLD
strWP_JPEG_FONTBOLD = WP_JPEG_FONTBOLD
Dim strWP_JPEG_FONTSIZE
strWP_JPEG_FONTSIZE = WP_JPEG_FONTSIZE
Dim strWP_JPEG_FONTCOLOR
strWP_JPEG_FONTCOLOR = WP_JPEG_FONTCOLOR
Dim strWP_WATERMARK_TEXT
strWP_WATERMARK_TEXT = WP_WATERMARK_TEXT
Dim strWP_ALBUM_NAME
strWP_ALBUM_NAME = WP_ALBUM_NAME
Dim strWP_WATERMARK_WIDTH_POSITION
strWP_WATERMARK_WIDTH_POSITION = WP_WATERMARK_WIDTH_POSITION
Dim strWP_WATERMARK_HEIGHT_POSITION
strWP_WATERMARK_HEIGHT_POSITION = WP_WATERMARK_HEIGHT_POSITION
Dim numWP_SMALL_WIDTH
numWP_SMALL_WIDTH = WP_SMALL_WIDTH
Dim numWP_SMALL_HEIGHT
numWP_SMALL_HEIGHT = WP_SMALL_HEIGHT
Dim numWP_LIST_WIDTH
numWP_LIST_WIDTH = WP_LIST_WIDTH
Dim numWP_LIST_HEIGHT
numWP_LIST_HEIGHT = WP_LIST_HEIGHT
Dim strWP_WATERMARK_LOGO
strWP_WATERMARK_LOGO = WP_WATERMARK_LOGO
Dim strWP_WATERMARK_ALPHA
strWP_WATERMARK_ALPHA = WP_WATERMARK_ALPHA
Dim numWP_SMALL_PAGERCOUNT
numWP_SMALL_PAGERCOUNT = WP_SMALL_PAGERCOUNT
Dim numWP_LIST_PAGERCOUNT
numWP_LIST_PAGERCOUNT = WP_LIST_PAGERCOUNT
Dim strWP_SUB_DOMAIN
strWP_SUB_DOMAIN = WP_SUB_DOMAIN
Dim strWP_ALBUM_INTRO
strWP_ALBUM_INTRO = WP_ALBUM_INTRO
Dim strWP_UPLOAD_RENAME
strWP_UPLOAD_RENAME = WP_UPLOAD_RENAME
Dim strWP_WATERMARK_AUTO
strWP_WATERMARK_AUTO = WP_WATERMARK_AUTO
Dim numWP_INDEX_PAGERCOUNT
numWP_INDEX_PAGERCOUNT = WP_INDEX_PAGERCOUNT
Dim numWP_BLOGPHOTO_ID
numWP_BLOGPHOTO_ID = WP_BLOGPHOTO_ID
Dim strWP_HIDE_DIVFILESND
strWP_HIDE_DIVFILESND = WP_HIDE_DIVFILESND%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div class="ShowBlogHint"><%Call GetBlogHint()%></div>
	<div class="divHeader">WindsPhoto 系统设置</div>
		<div class="SubMenu">
			<a href="<%=ZC_BLOG_HOST%>zb_users/plugin/windsphoto/admin_main.asp"><span class="m-left">相册管理</span></a>
			<a href="<%=ZC_BLOG_HOST%>zb_users/plugin/windsphoto/admin_addtype.asp"><span class="m-left">新建相册</span></a>
			<a href="<%=ZC_BLOG_HOST%>zb_users/plugin/windsphoto/admin_setting.asp"><span class="m-left m-now">系统设置</span></a>	
			<a href="<%=ZC_BLOG_HOST%>zb_system/cmd.asp?act=pluginMng"><span class="m-right">退出</span></a>
			<a href="<%=ZC_BLOG_HOST%>zb_users/plugin/windsphoto/help.asp"><span class="m-right">帮助说明</span></a>
			<a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp#more"><span class="m-right">更多功能</span></a>
		</div>
<form name="edit" method="post" action="admin_savesetting.asp">
<div id="divMain2">
<div class="content-box"><!-- Start Content Box -->
<div class="content-box-header">
<ul class="content-box-tabs">
	<li><a href="#fragment-1" class="default-tab"><span>参数设置</span></a></li>
	<li><a href="#fragment-2"><span>水印设置</span></a></li>
</ul>
<div class="clear"></div></div>
<div class="content-box-content">
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
</div></div>
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