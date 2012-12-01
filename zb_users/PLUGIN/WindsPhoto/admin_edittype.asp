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

<%
Call System_Initialize
Call WindsPhoto_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

BlogTitle = "管 理 相 册"

%><!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">
	<div class="divHeader">WindsPhoto 修改相册属性</div>
    <div class="SubMenu">
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_main.asp"><span class="m-left m-now">相册管理</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_addtype.asp"><span class="m-left">新建相册</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_setting.asp"><span class="m-left">系统设置</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_system/admin/admin.asp?act=PlugInMng"><span class="m-right" >退出</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp"><span class="m-right" >帮助说明</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp#more"><span class="m-right" >更多功能</span></a>
    </div>

    <div id="divMain2">
    <%
    If IsNumeric(Request.QueryString("typeid")) = FALSE Then
        Call SetBlogHint_Custom("!! 参数错误.")
        Response.Redirect"admin_main.asp"
    Else
        typeid = CInt(Request.QueryString("typeid"))
    End If

    sql = "SELECT * FROM WindsPhoto_zhuanti where id="&typeid
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, objConn, 1, 1
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
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->