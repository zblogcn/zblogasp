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
<% 'On Error Resume Next %>
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
%>
<!-- #include file="data/conn.asp" -->
<%
'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

BlogTitle = "新 建 相 册"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain">
	<div class="divHeader">WindsPhoto 新建相册</div>
    <div class="SubMenu">
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_main.asp"><span class="m-left">相册管理</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_addtype.asp"><span class="m-left m-now">新建相册</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_setting.asp"><span class="m-left">系统设置</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_system/admin/admin.asp?act=PlugInMng"><span class="m-right" >退出</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp"><span class="m-right" >帮助说明</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp#more"><span class="m-right" >更多功能</span></a>
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
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->