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
Call System_Initialize
Call WindsPhoto_Initialize()
%><%
'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

BlogTitle = "管 理 相 册"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
	<div class="divHeader">WindsPhoto 后台首页</div>
    <div class="SubMenu">
<script type="text/javascript">ActiveLeftMenu("aWindsPhoto")</script>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_main.asp"><span class="m-left m-now">相册管理</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_addtype.asp"><span class="m-left">新建相册</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/admin_setting.asp"><span class="m-left">系统设置</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_system/admin/admin.asp?act=PlugInMng"><span class="m-right" >退出</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp"><span class="m-right" >帮助说明</span></a>
  <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/windsphoto/help.asp#more"><span class="m-right" >更多功能</span></a>
    </div>

    <div id="divMain2">
    
    <%
    Dim ipagecount
    Dim ipagecurrent
    Dim irecordsshown
    If request.querystring("page") = "" Then
        ipagecurrent = 1
    Else
        ipagecurrent = CInt(request.querystring("page"))
    End If
    Set rso = Server.CreateObject("ADODB.RecordSet")
    sql = "select * FROM WindsPhoto_zhuanti order by ordered,id asc"
    sql2 = "select count(*) as C FROM WindsPhoto_desktop"
    rso.Open sql2, objConn, 3, 3
    sm = rso("c")
    rso.Close
    Set rso = Nothing
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.pagesize = 18
    rs.Open sql, objConn, 1, 1
    ipagecount = rs.pagecount
    If ipagecurrent > ipagecount Then ipagecurrent = ipagecount
    If ipagecurrent < 1 Then ipagecurrent = 1
    If ipagecount = 0 Then
        response.Write "<p align='center'>还没有任何相册>>><a href='admin_addtype.asp'>赶快添加一个新相册</a></p>"
    Else
        rs.absolutepage = ipagecurrent
        irecordsshown = 0
        Response.Write"<p>截止到 "&Now()&" 共有"&rs.RecordCount&"个相册,"&sm&"张图片 <a href='admin_html.asp'><font color='red'>生成静态首页/缓存文件↓</font></a></p>"
        Response.Write"<table width='100%' border='0' cellspacing='0' cellpadding='5'>"
        Do While irecordsshown<18 And Not rs.EOF
            Response.Write"<tr align='center'>"

            For i = 1 To 3

                If Not rs.EOF Then
                    sqlp = "select top 1 * FROM WindsPhoto_desktop where zhuanti="&rs("id")&" and hot<>0 order by id asc"
                    Set rsp = Server.CreateObject("ADODB.RecordSet")
                    rsp.Open sqlp, objConn, 1, 1
                    If rsp.EOF Or rsp.bof Then
                        surl = "images/notop.gif"
                    Else
                        surl = rsp("surl")
                        If InStr(surl, "photo.163.com") Or InStr(surl, "126.net") Or InStr(surl, "photo.sina.com") Or InStr(surl, "photos.baidu.com") Then surl = WP_SUB_DOMAIN &"stealink.asp?" & surl End If
                    End If
                    rsp.Close
                    Set rsp = Nothing
                    Set rso = Server.CreateObject("ADODB.RecordSet")
                    sql = "select count(*) as C FROM WindsPhoto_desktop where zhuanti="&rs("id")&""
                    rso.Open sql, objConn, 3, 3
                    sm = rso("c")
                    rso.Close
                    Set rso = Nothing
                    Dim sqlp
                    Set rsp = Server.CreateObject("ADODB.RecordSet")
                    sqlp = "select pass FROM WindsPhoto_zhuanti where id="&rs("id")&""
                    rsp.Open sqlp, objConn, 1, 1
                    p = rsp("pass")
                    If p<>"" Then
                        surl = "images/nopass.gif"
                    End If
                    rsp.Close
                    Set rsp = Nothing
                    Response.Write"<td width='33%'><a href='admin_addphoto.asp?typeid="&rs("id")&"'><img class='wp_top' src='"&surl&"' onload='WindsPhotoResizeImage(this,"&WP_SMALL_WIDTH&","&WP_SMALL_HEIGHT&")' /></a><p>["&rs("ordered")&"]<a href='"&WP_SUB_DOMAIN&"album.asp?typeid="&rs("id")&"' target='_blank'>"&rs("name")&"</a> | <a href='admin_edittype.asp?typeid="&rs("id")&"'>修改/删除</a> | <a href='admin_addphoto.asp?typeid="&rs("id")&"'><font color=red>上传/管理</font></a></p></td>"
                    irecordsshown = irecordsshown + 1
                    rs.movenext
                End If

            Next

            Response.Write"</tr>"
        Loop

        Response.Write"</table>"

        If ipagecount >1 then
            'Response.Write"<p>"&ipagecount&"页中的第"&ipagecurrent&"页 "
            Response.Write"<p><a title='首页' href='?&page=1'>[1]</a>"
            If ipagecurrent=1 then
                Response.Write"[上一页]"
            Else
                Response.Write"<a href='?page="&ipagecurrent-1&"'>[上一页]</a>"
            End If

            If ipagecount>ZC_PAGEBAR_COUNT Then
                a=ipagecurrent-Cint((ZC_PAGEBAR_COUNT-1)/2)
                b=ipagecurrent+ZC_PAGEBAR_COUNT-Cint((ZC_PAGEBAR_COUNT-1)/2)-1
                If a<=1 Then
                    a=1:b=ZC_PAGEBAR_COUNT
                End If
                If b>=ipagecount Then
                    b=ipagecount:a=ipagecount-ZC_PAGEBAR_COUNT+1
                End If
            Else
                a=1:b=ipagecount
            End If
            
            For i = a to b
                'ipagenow = ipagenow + 1
                If ipagecurrent = i Then
                    Response.Write"<span class=""now-page"">["&i&"]</span>"
                Else
                    Response.Write"<a href='?page="&i&"'>["&i&"]</a>"
                End If
            Next

            If ipagecount>ipagecurrent then
                Response.Write"<a href='?page="&ipagecurrent+1&"'>[下一页]</a>"
            Else
                Response.Write"[下一页]"
            End If

            Response.Write"<a title='尾页' href='?page="&ipagecount&"'>["&ipagecount&"]</a>"

            Response.Write"</p>"
        End If

    End If
    rs.Close
    Set rs = Nothing
    %>
    </div>
    <br><br><p align=center>Plugin Powered by <a href="http://www.wilf.cn" target="_blank">Wilf.cn</a></p>
</div>

<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->