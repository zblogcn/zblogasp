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
<!-- #include file="../../function/c_function.asp" -->
<!-- #include file="../../function/c_system_lib.asp" -->
<!-- #include file="../../function/c_system_base.asp" -->
<!-- #include file="../../function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<!-- #include file="data/conn.asp" -->
<%
Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

BlogTitle = "管 理 相 册"

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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
	<div class="Header">WindsPhoto 修改相册属性</div>
    <div class="SubMenu">
        <span class="m-left m-now"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/admin_main.asp">相册管理</a></span>
        <span class="m-left"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/admin_addtype.asp">新建相册</a></span>
        <span class="m-left"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/admin_setting.asp">系统设置</a></span>
        <span class="m-right"><a href="<%=ZC_BLOG_HOST%>cmd.asp?act=pluginMng">退出</a></span>
        <span class="m-right"><a href="<%=ZC_BLOG_HOST%>plugin/windsphoto/help.asp">帮助说明</a></span>
        <span class="m-right"><a href="<%=ZC_BLOG_HOST%>PLUGIN/windsphoto/help.asp#more">更多功能</a></span>
    </div>

    <div id="divMain2">
    <%
    If IsNumeric(Request.QueryString("typeid")) = FALSE Then
        Call SetBlogHint_Custom("!! 参数错误.")
        Response.Redirect"admin_main.asp"
    Else
        typeid = CInt(Request.QueryString("typeid"))
    End If

    sql = "SELECT * FROM zhuanti where id="&typeid
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, Conn, 1, 1
    If rs.EOF And rs.bof Then
        Call SetBlogHint_Custom("!! 还没有该相册.")
        Response.Redirect"admin_main.asp"
    Else
    %>
    <form id="edit" name="form3" method="post" action="admin_savetype.asp">
    <p>相册名称:<input type="text" name="name" size="50" maxlength="50" value="<%=rs("name")%>"></p>
    <p>相册排序:<input type="text" name="ordered" size="20" maxlength="20" value="<%=rs("ordered")%>"></p>
    <p>发布日期:<input type="text" name="fabu" size="20" maxlength="50" value="<%=rs("time1")%>"> 拍摄日期:<input type="text" name="riqi" size="20" maxlength="50" value="<%=rs("data")%>"></p>
    <p>查看密码:<input type="text" name="pass" size="20" maxlength="50" value="<%=rs("pass")%>"> 留空则任何人都可以浏览,输入 no 则不显示</p>
    <p>显示方式:<input type="radio" name="view" value="0" <%if rs("view") =false then response.write "checked" end if%>>  缩略图 <input type="radio" name="view" value="1" <%if rs("view") =true then response.write "checked" end if%>>列表</p>
    <p>相关介绍:<br><textarea name="js" rows=7 id="mce_editor_2" style="width: 435px;"><%=rs("js")%></textarea></p>
    <p>操作选项:<input type="radio" name="zt" value="editzhuanti" checked>修改 <input type="radio" name="zt" onClick="alert(alt)" alt="你选中了删除，该操作会删除整个相册，且不可逆，请慎重对待！" value="delzhuanti">删除</p>
    <input type="hidden" name="typeid" value="<%=typeid%>">
    <p>
    <input type="submit" name="submit" class="button" value="确定">
    <input type="reset" name="reset" class="button" value="重置">
    </p>
    </form>
    <%
    End If
    %>
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