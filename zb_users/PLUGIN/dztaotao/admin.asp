<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.8
'// 插件制作:    
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%
'On Error Resume Next
 %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_event.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_manage.asp" -->
<!-- #include file="../p_config.asp" -->

<% 
Call System_Initialize() 

'检查非法链接
Call CheckReference("") 

'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 

If CheckPluginState("dztaotao")=False Then Call ShowError(48)

BlogTitle="dztaotao - 查看/操作淘淘" 
%>

<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
	<div id="divMain">
		<div class="divHeader"><%=BlogTitle%></div>
        <div id="ShowBlogHint"><%Call GetBlogHint()%></div>
			<div class="SubMenu">
				<a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/admin.asp?a=list&page=1"><span class="m-left m-now">淘淘管理</span></a>
                <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/admin_cmt.asp?a=list&page=1"><span class="m-left">评论管理</span></a>
                <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/setting.asp"><span class="m-left">配置管理</span></a>
				<a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/help.asp"><span class="m-left">帮助说明</span></a>
			</div>
	<div id="divMain2">

<%
dim a
a=request.QueryString("a")
select case a
case "list"
%>
<table border="1" width="100%" cellpadding="2" cellspacing="0" bordercolordark="#f7f7f7" bordercolorlight="#cccccc">
<tr>
<td width="3%" bgcolor="#f7f7f7"><div align="center">ID</div></td>
<td width="8%" bgcolor="#f7f7f7"><div align="center">昵称</div></td>
<td width="8%" bgcolor="#f7f7f7"><div align="center">博客</div></td>
<td width="46%" bgcolor="#f7f7f7"><div align="center">说说内容</div></td>
<td width="5%" bgcolor="#f7f7f7"><div align="center">评论</div></td>
<td width="13%" bgcolor="#f7f7f7"><div align="center">说的时间</div></td>
<td width="17%" bgcolor="#f7f7f7"><div align="center">操作</div></td>
</tr>
<%
Dim objRS,page
dim r_rs,t
dim r_recordcount
page = Request.Querystring("Page")
Set objRS=objConn.Execute("SELECT * FROM [dz_taotao] ORDER BY [id] desc")
If (Not objRS.bof) And (Not objRS.eof) Then
	Const MaxPerPage=15
	Dim CurrentPage,F
	Dim TotalPut
	objRS.MoveFirst
	If Trim(page)<>"" Or Not IsNumeric(page) Then
	CurrentPage=Clng(page)
	Else
	CurrentPage=1
	End If
	TotalPut=objConn.ExeCute("Select Count(id) From dz_taotao",0,1)(0)
	If CurrentPage<>1 Then
		If (CurrentPage-1)*MaxPerPage<TotalPut Then
		objRS.Move(CurrentPage-1)*MaxPerPage
		End If
	End If
	Dim N,K
	If (TotalPut Mod MaxPerPage)=0 Then
	N=TotalPut \ MaxPerPage
	Else
	N=TotalPut \ MaxPerPage+1
	End If
	For F=1 To MaxPerPage
	If Not objRS.Eof Then
t=t+1
%>
<tr>
<td><%=objRS("id")%></td>
<td><%=objRS("username")%></td>
<td><%=objRS("site")%></td>
<td><textarea name="<%=t%>" id="<%=t%>" cols="28" rows="2" wrap="VIRTUAL" class="inputt" style="width:430px;"><%=objRS("content")%></textarea></td>
<td>&nbsp;<%=objRS("comments")%>条</td>
<td><%=objRS("addtime")%></td>
<td><% if objRS("itype") = 0 then %><a href="admin.asp?a=nochk&id=<%=objRS("id")%>&page=<%=page%>">已审核</a><%else%><a href="admin.asp?a=ischk&id=<%=objRS("id")%>&page=<%=page%>" style="color:#F00;">未审核</a><%end if%>   <a href="admin.asp?a=update_cmt&id=<%=objRS("id")%>&page=<%=page%>" title="更新评论">更新</a>  <a href="admin.asp?a=r&id=<%=objRS("id")%>">查看评论</a>  <a href="admin.asp?page=<%=page%>&a=del&id=<%=objRS("id")%>" onClick="return confirm('您真的要删除该说说吗？');">删除</a></td>
</tr>
<%
	objRS.MoveNext
	End If
    Next
    Else
		response.write "<tr><td colspan='7'>暂时数据</td></tr>"
End If
objRS.Close
Set objRS=Nothing

K=CurrentPage

response.write "<tr><td colspan='7'><div class=""pagebar"">"&ExportPageBar(page,n,MaxPerPage,"admin.asp?a=list&page=")&"</div></td></tr>"

%>
</table>
<%
case "update_cmt"'更新评论数量

dim t_cmt_count
	id=request.QueryString("id")
	page=request.QueryString("page")
	if not isnumeric(id) then
	response.write "没有找到您要操作的信息"
	response.End()
	end if
	set t_cmt_count = objConn.execute("select count(*) as c from [dz_comment] where tt_id="&id&" and itype=0")
	if not t_cmt_count.eof then
	objConn.execute("update [dz_taotao] set comments = "&t_cmt_count("c")&" where id="&id&"")
	end if
	response.Redirect("admin.asp?a=list&page="&page)
	response.End()


case "s_sava"
isput=request.form("isput")
page_count=trim(request.form("page_count"))

    if isput="" then
  	response.write "<script language=javascript>"	
		response.write "alert('显示方式不能为空');"	
		response.write "</script>"
		response.write "<script language=javascript>location='javascript:history.back(1)'</script>"
   Response.End
   end if

    if page_count="" then
  response.write "<script language=javascript>"	
		response.write "alert('每页显示多少条啊？');"	
		response.write "</script>"
		response.write "<script language=javascript>location='javascript:history.back(1)'</script>"
   Response.End
   end if
   
   objConn.execute("update dz_set set isput='"&isput&"',page_count="&page_count&"")
   response.write "更新成功<br><a href='?a=s'>点击返回</a>"

'配置说说
case "s"
s="select * from dz_set"
set r=server.createobject("ADODB.RecordSet")
r.open s,objConn,1,3
if not r.eof then
isput=r("isput")
page_count=r("page_count")
end if
r.close:set r=nothing
%>
<form action="admin.asp?a=s_sava" method="post" id="form2" name="form2">
<table border="1" width="98%" cellpadding="2" cellspacing="0" bordercolordark="#f7f7f7" bordercolorlight="#cccccc">
<tr>
  <td colspan="4"><strong>基本设置</strong></td>
  </tr>
<tr>
<td width="10%">前台发布：</td>
<td colspan="3" width="90%"><input name="isput" type="radio" id="is_isput" value="0"<%if isput=0 then response.write " checked"%> />允许
<input name="isput" type="radio" id="no_isput" value="1"<%if isput=1 then response.write " checked"%> />不允许
</td>
</tr>
<tr>
  <td width="10%">每页显示：</td>
  <td width="90%">
	<input name="page_count" type="text" id="page_count" value="<%=page_count%>" size="4" />条
  </td>
</tr>
<tr>
  <td colspan="2"><strong>广告设置</strong></td>
  </tr>
<tr>
  <td>顶部广告1：</td>
  <td><table width="100%" border="1">
    <tr>
      <td width="9%">图片：</td>
      <td width="91%"><input name="textfield3" type="text" id="textfield3" size="60" /></td>
    </tr>
    <tr>
      <td>链接：</td>
      <td><input name="textfield4" type="text" id="textfield4" size="60" /></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td>顶部广告2：</td>
  <td><table width="100%" border="1">
    <tr>
      <td width="9%">图片：</td>
      <td width="91%"><input name="textfield5" type="text" id="textfield5" size="60" /></td>
    </tr>
    <tr>
      <td>链接：</td>
      <td><input name="textfield5" type="text" id="textfield6" size="60" /></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td>顶部广告3：</td>
  <td><table width="100%" border="1">
    <tr>
      <td width="9%">图片：</td>
      <td width="91%"><input name="textfield6" type="text" id="textfield7" size="60" /></td>
    </tr>
    <tr>
      <td>链接：</td>
      <td><input name="textfield6" type="text" id="textfield8" size="60" /></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td>顶部广告4：</td>
  <td><table width="100%" border="1">
    <tr>
      <td width="9%">图片：</td>
      <td width="91%"><input name="textfield7" type="text" id="textfield9" size="60" /></td>
    </tr>
    <tr>
      <td>链接：</td>
      <td><input name="textfield7" type="text" id="textfield10" size="60" /></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td colspan="2"><strong>内容广告</strong></td>
  </tr>
<tr>
  <td>显示位置：</td>
  <td>在第
    <input name="textfield" type="text" id="textfield" size="4" />
    条淘淘之后显示广告</td>
</tr>
<tr>
  <td>广告代码：</td>
  <td><textarea name="textfield2" cols="60" rows="4" id="textfield2"></textarea></td>
</tr>
<tr>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
</tr>
<tr>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
</tr>
<tr>
  <td>&nbsp;</td>
  <td><input type="submit" name="button3" id="button3" value="提交" /></td>
</tr>

</table>
</form>
<%
'保存提交的说说
case "s_post"
dim username,site,content
username=request.form("username")
site=request.form("site")
content=request.form("content")
if username<>"" and content<>"" then
objConn.execute("insert into dz_taotao (username,site,content) values ('"&username&"','"&site&"','"&content&"')")
response.write "添加成功<br><a href='admin.asp?a=list'>返回列表</a>"
response.end
else
response.write "<script>alert('能填的都要填啊！');history.back();</script>"
response.end
end if

case "p"
%>
<form id="pform1" name="pform1" method="post" action="?a=s_post">
<table border="1" width="98%" cellpadding="2" cellspacing="0" bordercolordark="#f7f7f7" bordercolorlight="#cccccc" class="css">
<tr>
  <td colspan="4" bgcolor="#f7f7f7">发布内容</td>
  </tr>
<tr>
  <td width="10%" bgcolor="#f7f7f7">昵称：</td>
  <td width="90%" colspan="3" bgcolor="#f7f7f7"><input name="username" type="text" id="username" maxlength="30" /></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">博客：</td>
  <td colspan="3" bgcolor="#f7f7f7"><input name="site" type="text" id="site" maxlength="200" /></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">内容：</td>
  <td colspan="3" bgcolor="#f7f7f7"><textarea name="content" cols="55" rows="5" id="content"></textarea></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">&nbsp;</td>
  <td colspan="3" bgcolor="#f7f7f7"><input type="submit" name="button" id="button" value="增加" />
    <input type="button" name="button2" id="button2" value="返回列表" onClick="location.href='?a=list';" /></td>
</tr>
</table>

</form>
<br><br><br>
<%
case "nochk"'已审核，点击后改成未审核
dim id
id=request.QueryString("id")
page=request.QueryString("page")
if not isnumeric(id) then
response.write "没有找到您要查看的信息"
response.End()
end if
objConn.execute("update [dz_taotao] set itype = 4 where id="&id&"")
response.Redirect("admin.asp?a=list&page="&page)
response.End()

case "ischk"'未审核，点击后改成已审核
id=request.QueryString("id")
page=request.QueryString("page")
if not isnumeric(id) then
response.write "没有找到您要查看的信息"
response.End()
end if
objConn.execute("update [dz_taotao] set itype = 0 where id="&id&"")
response.Redirect("admin.asp?a=list&page="&page)
response.End()



case "del_r"
id=request.QueryString("id")
tt_id=request.QueryString("tt_id")
if not isnumeric(id) then
response.write "没有找到您要查看的评论"
response.End()
end if
objConn.execute("delete from dz_comment where id="&id&"")
response.Redirect("admin.asp?a=r&id="&tt_id)
response.End()

case "r"
id=request.QueryString("id")
if not isnumeric(id) then
response.write "没有找到您要查看的评论"
response.End()
end if
%>
<table border="1" width="98%" cellpadding="2" cellspacing="0" bordercolordark="#f7f7f7" bordercolorlight="#cccccc" class="css">
<tr>
  <td colspan="4" bgcolor="#f7f7f7"><b>查看评论</b></td>
  </tr>
<%
dim rs
set rs=objConn.execute("select * from dz_comment where tt_id="&id&"")
if not rs.eof then
do while not rs.eof
%>
<tr>
  <td width="10%" bgcolor="#f7f7f7">时间：</td>
  <td width="90%" colspan="3" bgcolor="#f7f7f7"><%=rs("addtime")%></td>
</tr>

<tr>
  <td width="10%" bgcolor="#f7f7f7">昵称：</td>
  <td width="90%" colspan="3" bgcolor="#f7f7f7"><%=rs("u_sername")%></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">博客：</td>
  <td colspan="3" bgcolor="#f7f7f7"><%=rs("u_site")%></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">评论内容：</td>
  <td colspan="3" bgcolor="#f7f7f7"><%=rs("content")%></td>
</tr>
<tr>
  <td bgcolor="#f7f7f7">&nbsp;</td>
  <td colspan="3" bgcolor="#f7f7f7"><a href="javascript:history.back();">返回列表页</a>   <a href="admin.asp?a=del_r&id=<%=rs("id")%>&tt_id=<%=id%>">删除</a>
    
    </td>
</tr>
<tr>
  <td colspan="4" style="height:1px; background:#CCC;"></td>
</tr>


<%
rs.movenext
loop
end if
rs.close:set rs=nothing
%>
</table>
<%
case "del"
id=request.QueryString("id")
page=request.QueryString("page")
if not isnumeric(id) then
response.write "<script>alert('操作提示：对不起，没有找到您要删除的信息！');history.back();</script>"
response.End()
end if
objConn.execute("delete from dz_comment where tt_id="&id&"")
objConn.execute("delete from dz_taotao where id="&id&"")
response.Redirect("admin.asp?a=list&page="&page)
response.End()

end select
%>
</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>

