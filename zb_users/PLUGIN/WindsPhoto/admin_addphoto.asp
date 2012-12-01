<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8 其它版本未知
'// 插件制作:    狼的旋律(http://www.wilf.cn) / winds(http://www.lijian.net)
'// 备    注:    WindsPhoto
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

<%Call System_Initialize
Call WindsPhoto_Initialize()%><%

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>2 Then Call ShowError(6)
If CheckpluginState("windsphoto") = FALSE Then Call ShowError(48)

BlogTitle = "WindsPhoto 上传/管理"

%>
<%
action = Request.QueryString("action")
typeid = Request.QueryString("typeid")
%>

<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
	<link rel="stylesheet" rev="stylesheet" href="images/windsphoto.css" type="text/css" media="screen" />

	<script type="text/javascript" src="script/windsphoto.js"></script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->


<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
	<div class="divHeader">WindsPhoto <%if action = "insert" then%>点击图片插入<%else%>上传/管理<%end if%></div>
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
if action = "" then

If IsNumeric(Request.QueryString("typeid")) = FALSE Then
    Call SetBlogHint_Custom("!! 参数错误.")
    Response.Redirect"admin_main.asp"
Else
    typeid = CInt(Request.QueryString("typeid"))
End If
sql1 = "SELECT * FROM WindsPhoto_zhuanti where id="&typeid
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.Open sql1, objConn, 1, 1
If rs1.EOF And rs1.bof Then
    Call SetBlogHint_Custom("!! 还没有该相册.")
    Response.Redirect"admin_main.asp"
Else
    Response.Write"<p><a href='"&WP_SUB_DOMAIN&"default.asp' target='_blank'>"&WP_ALBUM_NAME&"</a> &raquo; <a href='"&WP_SUB_DOMAIN&"album.asp?typeid="&rs1("id")&"' target='_blank'>"&rs1("name")&"</a> <a href="""" onclick=""if(document.getElementById('moreupload').style.display=='none'){document.getElementById('localupload').style.display='block';document.getElementById('moreupload').style.display='block';document.getElementById('remoteupload').style.display='none';}else{document.getElementById('moreupload').style.display='none'};return false;"" ><font color=red>批量上传↑</font></a> <a href="""" onclick=""if(document.getElementById('remoteupload').style.display=='none'){document.getElementById('remoteupload').style.display='block';document.getElementById('localupload').style.display='none';document.getElementById('moreupload').style.display='none';}else{document.getElementById('remoteupload').style.display='none';document.getElementById('localupload').style.display='block';};return false;"" ><font color=green>远程图片↑</font></a></p>"
    Response.Write"<form id=""edit"" name=""windsphoto"" action=""admin_uploadpic.asp"" method=""post"" enctype=""multipart/form-data"" onsubmit=""return CheckForm()"">"
    Response.Write"<input type=""hidden"" name=""zhuanti"" id=""zhuanti"" value="""&rs1("id")&""">"
	Response.Write"<input type=""hidden"" name=""category"" id=""category"" value="""&rs1("name")&""">"
End If

rs1.Close
Set rs1 = Nothing
%>
<p>
	标题:<input type="text" name="name" id="name" maxlength="50" size="30" /> <%if WP_IF_ASPJPEG="1" then%>水印<input type="checkbox" name="mark" id="mark" <%if WP_WATERMARK_AUTO="1" then%>checked<%end if%> value="1" /><%end if%> 文件重命名<input type="checkbox" name="autoname" id="autoname" <%if WP_UPLOAD_RENAME="1" then%>checked<%end if%> value="1" />
</p>
<div style="display:none;" id="remoteupload">
	<p>地址:<input type="text" name="url" maxlength="120" size="30" /> <font color="red">(*)必填</font></p>
	<p>缩图:<input type="text" name="surl" maxlength="120" size="30" /> </p>
</div>
<div style="display:block;" id="localupload">
	<p>上传:<input type="file" name="file0" size="30" /> <font color="red">(*)必填</font></p>
</div>
<div style="display:none;" id="moreupload">
	<p>更多:<input type="file" name="file1" size="30" /></p>
	<p>更多:<input type="file" name="file2" size="30" /> 建议3张，多了容易出错</p>
	<p>更多:<input type="file" name="file3" size="30" /></p>
	<p>更多:<input type="file" name="file4" size="30" /></p>
</div>
<p>照片简介:<br>
	<textarea name="photointro" id="mce_editor_2" style="width: 400px;height:100px;"></textarea>
</p>
<p>
	<input type="submit" id="upupup" value="提交" name="submit" class="button" />
	<input type="reset" id="reset" value="重置" name="reset" class="button" />
	<input type="hidden" name="quick" value="0" />
	<input type="hidden" name="act" value="upload" />
</p>
</form>
<p><a name="exist"></a>已有图片:</p>
<%else%>
<script language="JavaScript">
function jumpto()
{
location.href=document.category.list.value;
}
</script>
<form name="category">
<select name="list" onChange="jumpto()">
<option>---选择相册分类---</option>
<%
sql2 = "SELECT * FROM WindsPhoto_zhuanti order by ordered,id asc"
Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.Open sql2, objConn, 1, 1
If rs2.EOF And rs2.bof Then
	Response.Write"<option>还没有相册</option>"
Else
	Do While Not rs2.EOF
		Response.Write"<option value='admin_addphoto.asp?typeid="&rs2("id")&"&action=insert'>"&rs2("name")&"</option>"
		rs2.movenext
	Loop
End If
rs2.Close
Set rs2 = Nothing
%>
</select>
</form>
<%end if%>
<%
Response.Write"<table width='100%' border='0' cellspacing='0' align='center'>"

Dim ipagecount
Dim ipagecurrent
Dim irecordsshown
If request.querystring("page") = "" Then
    ipagecurrent = 1
Else
    ipagecurrent = CInt(request.querystring("page"))
End If

If WP_ORDER_BY="0" then
    sql = "SELECT * FROM WindsPhoto_desktop where zhuanti="&typeid&" ORDER BY id asc"
Else
    sql = "SELECT * FROM WindsPhoto_desktop where zhuanti="&typeid&" ORDER BY id desc"
End If

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, objConn, 1, 1
rs.pagesize = 24
ipagecount = rs.pagecount
If ipagecurrent > ipagecount Then ipagecurrent = ipagecount
If ipagecurrent < 1 Then ipagecurrent = 1
If ipagecount = 0 Then
    response.Write "<tr><td align='center'><img src='images/nopic.jpg'></td></tr></table>"
Else
    Count = 1
    rs.absolutepage = ipagecurrent
    irecordsshown = 0
    Do While irecordsshown<24 and Not rs.EOF
	url = rs("url")
	surl = rs("surl")
	If InStr(surl, "photo.163.com") Or InStr(surl, "photo.sina.com") Or InStr(surl, "photos.baidu.com") Then surl = "stealink.asp?" & surl End If
htmlurl = "<img src="&ZC_BLOG_HOST&"zb_users/plugin/windsphoto/"&url&" />"
%>
<%
If(Count Mod 4 = 1) Then response.Write "<tr>"
    response.Write "<td align=center>"
%>
<%if action = "insert" then%><a href="javascript:try{window.opener.document.getElementById('MyEditor___Frame').contentWindow.frames[0].document.getElementsByTagName('body')[0].innerHTML+='<%=htmlurl%>'}catch(e){};window.close();"><%else%><a href="#"><%end if%><img class="wp_top" src="<%=surl%>" /></a><%if action <> "insert" then%><p align="center"><%if Rs("hot")="" OR IsNULL(Rs("hot"))=true or Rs("hot")=0 then%><a href='admin_savephoto.asp?action=hot&id=<%=rs("id")%>&typeid=<%=typeid%>&t=1'>封面</a><%else%><a href='admin_savephoto.asp?action=hot&id=<%=rs("id")%>&typeid=<%=typeid%>&t=0'><span style="color:red">封面</span></a><%end if%>|<a class="popup" href="admin_editphoto.asp?action=edit&id=<%=rs("id")%>">编辑</a>|<a onClick="Javascript:if(confirm('确定要删除吗?')){return true;}else{return false;}" href='admin_savephoto.asp?action=del&id=<%=rs("id")%>&typeid=<%=typeid%>'>删除</a>]</p><%end if%>
<%
response.Write "</td>"
If(Count Mod 4 = 0) Then
    response.Write "</tr>"
End If
Count = Count + 1
irecordsshown = irecordsshown + 1
rs.movenext
Loop

Response.Write"</table>"

    '分页
    if ipagecount >1 then
        Response.Write"<p>"&ipagecount&"页中的第"&ipagecurrent&"页 "
        if ipagecurrent=1 then
            Response.Write"[上一页] "
        else
            Response.Write"<a href='?typeid="&typeid&"&page="&ipagecurrent-1&"'>[上一页]</a> "
        end if

        for i = 1 to ipagecount
        ipagenow = ipagenow + 1
        if ipagecurrent=ipagenow then
            Response.Write"["&ipagenow&"] "
    else
    Response.Write"<a href='?typeid="&typeid&"&page="&ipagenow&"'>["&ipagenow&"]</a> "
        end if
    next

        if ipagecount>ipagecurrent then
            Response.Write"<a href='?typeid="&typeid&"&page="&ipagecurrent+1&"'>[下一页]</a> "
        else
            Response.Write"[下一页] "
        end if

        Response.Write"</p>"
    end if

End If
%>
    </div><br><br><p align=center>Plugin Powered by <a href="http://www.wilf.cn/" target="_blank">Wilf.cn</a></p>
</div>
</body>
</html>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->