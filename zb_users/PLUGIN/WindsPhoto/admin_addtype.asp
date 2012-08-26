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
<!-- #include file="../p_config.asp" -->
<!-- #include file="data/conn.asp" -->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

BlogTitle = "新 建 相 册"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta name="robots" content="noindex,nofollow"/>
	<link rel="stylesheet" rev="stylesheet" href="../../CSS/admin.css" type="text/css" media="screen" />
	<title><%=BlogTitle%></title>
</head>
<body>
<div id="divMain">
	<div class="Header">WindsPhoto 新建相册</div>
    <div class="SubMenu">
        <span class="m-left"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/admin_main.asp">相册管理</a></span>
        <span class="m-left m-now"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/admin_addtype.asp">新建相册</a></span>
        <span class="m-left"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/admin_setting.asp">系统设置</a></span>
        <span class="m-right"><a href="<%=ZC_BLOG_HOST%>cmd.asp?act=pluginMng">退出</a></span>
        <span class="m-right"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/help.asp">帮助说明</a></span>
        <span class="m-right"><a href="<%=ZC_BLOG_HOST%>PLUGIN/windsphoto/help.asp#more">更多功能</a></span>
    </div>

    <div id="divMain2">
        <form id="edit" name="form3" method="post" action="admin_savetype.asp?zt=addzhuanti">
        <p>相册名称:<input type="text" name="name" size="50" maxlength="50"> </p>
        <p>相册排序:<input type="text" name="ordered" size="20" maxlength="10"></p>
        <p>发布日期:<input type="text" name="fabu" size="20" maxlength="50" value="<%=date()%>"> 拍摄日期:<input type="text" name="riqi" size="20" maxlength="50" value="<%=date()%>"></p>
        <p>设置密码:<input type="text" name="pass" size="20" maxlength="50"> 留空则任何人都可以浏览,输入 no 则不显示</p>
        <p>显示方式:<input type="radio" name="view" value="0" checked> 缩略图 <input type="radio" name="view" value="1">列表</p>
        <p>相关介绍:<br><textarea name="js" rows="7" id="mce_editor_2" style="width: 435px;height:80px;"></textarea></p>
        <p><input type="submit" name="Submit3" class="button" value="确定"></p>
        </form>
    </div>
    <br><br><p align=center>Plugin Powered by <a href="http://www.wilf.cn" target="_blank">Wilf.cn</a></p>
</div>
</body>
</html>
<%
Call System_Terminate()

If Err.Number<>0 Then
    Call ShowError(0)
End If
%>