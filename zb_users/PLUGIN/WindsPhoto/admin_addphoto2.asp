<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备   注:    WindsPhoto
'// 最后修改：   2011.8.22
'// 最后版本:    2.7.3
'///////////////////////////////////////////////////////////////////////////////
%>
<% On Error Resume Next %>
<% Response.CodePage=65001%>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<!-- #include file="data/conn.asp" -->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

BlogTitle = "WindsPhoto 上传/管理"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta name="robots" content="noindex,nofollow"/>
	<link rel="stylesheet" rev="stylesheet" href="../../CSS/admin.css" type="text/css" media="screen" />
	<script language="JavaScript" src="../../script/common.js" type="text/javascript"></script>
</head>
<body>
    <form border="1" id="edit" name="upload" action="admin_uploadpic.asp" method="post" enctype="multipart/form-data" onsubmit="return CheckForm()">
    <input type="hidden" name="zhuanti" value="<%=WP_BLOGPHOTO_ID%>">
	<p>上传图片到WindsPhoto贴图相册:</p>
    <p>
    <input type="file" name="file0" size="20"><input type="hidden" name="time" value="<%=now()%>"> <input type="submit" id="upupup" value="提交" name="submit" class="button" /> <input type="reset" id="reset" value="重置" name="reset" class="button" />
     <%if WP_IF_ASPJPEG="1" then%><input type="checkbox" name="mark" id="mark" <%if WP_WATERMARK_AUTO="1" then%>checked<%end if%> value="1">水印<%end if%> <input type="checkbox" name="autoname" id="autoname" <%if WP_UPLOAD_RENAME="1" then%>checked<%end if%> value="1">自动命名上传文件
    <input type="hidden" name="quick" value="1">
    <input type="hidden" name="act" value="upload">
    </p>
    </form>
</body>
</html>
<%
Call System_Terminate()

If Err.Number<>0 Then
    Call ShowError(0)
End If
%>