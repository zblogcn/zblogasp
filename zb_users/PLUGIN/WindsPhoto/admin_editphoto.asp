<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备   注:     WindsPhoto
'// 最后修改：   2011.8.22
'// 最后版本:    2.7.3
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
Call System_Initialize()%><!-- #include file="data/conn.asp" --><%

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

BlogTitle = "管 理 相 册"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">
	<div class="divHeader">WindsPhoto 编辑图片</div>
    <div class="SubMenu">
        <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_main.asp"><span class="m-left m-now">相册管理</span></a>
        <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_addtype.asp"><span class="m-left">新建相册</span></a>
        <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_setting.asp"><span class="m-left">系统设置</span></a>
        <a href="<%=ZC_BLOG_HOST%>zb_system/cmd.asp?act=pluginMng"><span class="m-right">退出</span></a>
        <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp"><span class="m-right">帮助说明</span></a>
        <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp#more"><span class="m-right">更多功能</span></a>
    </div>

    <div id="divMain2">
    <%
    typen = Request("typeid")
    If Request.QueryString("action") = "edit" Then
        id = Request.QueryString("id")
        sql = "SELECT * FROM desktop where id="&id
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, Conn, 1, 1
        Name = rs("name")
        url = rs("url")
        surl = rs("surl")
        hot = rs("hot")
        zhuanti = rs("zhuanti")
        jj = rs("jj")
        previewsurl = rs("surl")
        itime = rs("itime")
        If InStr(previewsurl, "photo.163.com") Or InStr(previewsurl, "photo.sina.com") Or InStr(previewsurl, "photos.baidu.com") Then previewsurl = "stealink.asp?" & previewsurl End If
        rs.Close
        Set rs = Nothing

        title = "编辑文件"
        action = "edit"
    Else
        title = "添加文件"
        action = "addfile"
    End If
    %>
    <form id="edit" name="form" method="post" action="admin_savephoto.asp?action=<%=action%>&typeid=<%=typeid%>&id=<%=id%>">
    <p>标题:<input type="text" name="name" maxlength="30" value="<%=name%>" size="40"></p>
    <img src="<%=previewsurl%>" border="0" style="max-width:144px;float:right;" onerror="this.src='images/error.gif'" />
    <p>相册:
    <select name="zhuanti">
    <%
    sql = "SELECT * FROM zhuanti order by ordered,id asc"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.Open sql, Conn, 1, 1
    If rs1.EOF And rs1.bof Then
        Response.Write"<option>还没有相册</option>"
    Else
        Do While Not rs1.EOF
            Response.Write"<option value='"&rs1("id")&"'"
            If zhuanti = rs1("id") Then Response.Write" selected"
            Response.Write">"&rs1("name")&"</option>"
            rs1.movenext
        Loop
    End If
    rs1.Close
    Set rs1 = Nothing
    %>
    </select>
    </p>
    <p>封面:<input type=radio name="hot" value="1" <%If hot<>0 then response.write "checked" End If%>>是 <input type=radio name="hot" value="0" <%If hot=false then response.write "checked" End If%>>否</p>
    <p>时间:<input type=text name="itime" value="<%=itime%>" size="45"></p>
    <p>地址:<input type=text name="url" value="<%=url%>" size="45"></p>
    <p>缩图:<input type=text name="surl" value="<%=surl%>" size="45"></p>
    <p>照片简介:<br>
    <textarea name="jj" cols="50" rows="5" id="mce_editor_2"><%=jj%></textarea>
    </p>
    <p>
    <input type=submit value="确定" name="submit" class="button">
    <input type=reset value="重置" name="reset" class="button">
    </p>
    </form>
    </div>
    <br><br><p align=center>Plugin Powered by <a href="http://www.wilf.cn" target="_blank">Wilf.cn</a></p>
</div>
</body>
</html>
<%
Call System_Terminate()

'If Err.Number<>0 Then
 '   Call ShowError(0)
'End If
%>