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
		<div class="Header"><%=BlogTitle%></div>
        <div id="ShowBlogHint"><%Call GetBlogHint()%></div>
			<div class="SubMenu">
				<a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/admin.asp?a=list&page=1"><span class="m-left">淘淘管理</span></a>
                <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/admin_cmt.asp?a=list&page=1"><span class="m-left m-now">评论管理</span></a>
				<a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/admin.asp?a=p"><span class="m-left">发布说说</span></a>
                <a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/setting.asp"><span class="m-left">配置管理</span></a>
				<a href="<%=ZC_BLOG_HOST%>zb_users/PLUGIN/dztaotao/help.asp"><span class="m-left">帮助说明</span></a>
			</div>
	<div id="divMain2">
<%
dim a
a=request.QueryString("a")
select case a

case "updatelist"'更新列表时间
dim t_IDs,i
t_IDs = replace(request.Form("edtDel")," ","")
t_IDs = split(t_IDs,",")
for i=0 to Ubound(t_IDs)
objConn.execute("update [dz_taotao] set addtime = '"&request.form("s_time_"&t_IDs(i))&"' where id="&t_IDs(i)&"")
response.write request("s_time_"&t_IDs(i))&"<br>"
next

case "is_chk"'当前是已审核，点击后变成未审核
t=request.QueryString("t")
id=request.QueryString("id")
page=request.QueryString("page")
if not isnumeric(id) then
response.write "<script>alert('未找到您要操作的内容');</script>"
response.end 
end if
objConn.execute("update [dz_comment] set itype = -4 where id="&id&"")
	'更新评论
	set cc=objConn.execute("select count(*) as c_count from dz_comment where tt_id="&t&" and itype=0")
	if not cc.eof then
	objConn.execute("update [dz_taotao] set comments = "&cc("c_count")&" where id = "&t&"")
	end if
	cc.close:set cc=nothing

response.Redirect "admin_cmt.asp?a=list&page="&page

case "no_chk"'当前是未审核，点击后变成已审核
t=request.QueryString("t")
id=request.QueryString("id")
page=request.QueryString("page")
if not isnumeric(id) then
response.write "<script>alert('未找到您要操作的内容');</script>"
response.end 
end if
objConn.execute("update [dz_comment] set itype = 0 where id="&id&"")
	'更新评论
	set cc=objConn.execute("select count(*) as c_count from dz_comment where tt_id="&t&" and itype=0")
	if not cc.eof then
	objConn.execute("update [dz_taotao] set comments = "&cc("c_count")&" where id = "&t&"")
	end if
	cc.close:set cc=nothing

response.Redirect "admin_cmt.asp?a=list&page="&page

case "list"
%>
<form name="update_form1" id="update_form1" action="admin_cmt.asp?a=updatelist" method="post">
<table border="1" width="100%" cellpadding="2" cellspacing="0" bordercolordark="#f7f7f7" bordercolorlight="#cccccc">
<tr>
  <td width="3%" bgcolor="#f7f7f7"><a onClick="BatchSelectAll();return false" href="">全选</a></td>
<td width="3%" bgcolor="#f7f7f7"><div align="center">ID</div></td>
<td width="8%" bgcolor="#f7f7f7"><div align="center">昵称</div></td>
<td width="46%" bgcolor="#f7f7f7"><div align="center">评论内容</div></td>
<td width="13%" bgcolor="#f7f7f7"><div align="center">评论时间</div></td>
<td width="17%" bgcolor="#f7f7f7"><div align="center">操作</div></td>
</tr>
<%
Dim objRS,page
dim r_rs
dim r_recordcount
page = Request.Querystring("Page")
Set objRS=objConn.Execute("SELECT * FROM [dz_comment] ORDER BY [id] desc")
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
	TotalPut=objConn.ExeCute("Select Count(id) From dz_comment",0,1)(0)
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
  <td><input id="edtDel" type="checkbox" value="<%=objRS("id")%>" name="edtDel"></td>
<td><%=objRS("id")%></td>
<td><%=objRS("u_sername")%></td>
<td><textarea name="<%=t%>" id="<%=t%>" cols="28" rows="2" wrap="VIRTUAL" class="inputt" style="width:430px; height:30px;"><%=objRS("content")%></textarea></td>
<td><input type="text" name="s_time_<%=objRS("id")%>" id="s_time_<%=objRS("id")%>" value="<%=objRS("addtime")%>" /></td>
<td><%if objRS("itype")=0 then%><a href="admin_cmt.asp?a=is_chk&id=<%=objRS("id")%>&page=<%=page%>&t=<%=objRS("tt_id")%>" title="已审核">已审核</a><%else%><a href="admin_cmt.asp?a=no_chk&id=<%=objRS("id")%>&page=<%=page%>&t=<%=objRS("tt_id")%>" title="未审核" style="color:#F00">未审核</a><%end if%>  <a href="view.asp?id=<%=objRS("tt_id")%>" target="_blank">查看内容</a>  <a href="admin_cmt.asp?page=<%=page%>&a=del&id=<%=objRS("id")%>&t=<%=objRS("tt_id")%>">删除</a></td>
</tr>
<%
	objRS.MoveNext
	End If
    Next
    Else
		response.write "<tr><td colspan='6'>暂时数据</td></tr>"
End If
objRS.Close
Set objRS=Nothing

K=CurrentPage
response.write "<tr><td colspan='6'><div class=""pagebar"">"&ExportPageBar(page,n,MaxPerPage,"admin.asp?a=list&page=")&"</div></td></tr>"
%>

<tr><td colspan='6'><input type="submit" id="btnPost" value="更新" class="button"></td></tr>
</table>
</form>
<%
case "update_cmt"'更新评论数量


	id=request.QueryString("id")
	page=request.QueryString("page")
	if not isnumeric(id) then
	response.write "没有找到您要操作的信息"
	response.End()
	end if
	set t_cmt_count = objConn.execute("select count(*) as c from [dz_comment] where tt_id="&id&"")
	if not t_cmt_count.eof then
	objConn.execute("update [dz_taotao] set comments = "&t_cmt_count("c")&" where id="&id&"")
	end if
	response.Redirect("admin_cmt.asp")
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
<%
'保存提交的说说
case "s_post"
username=request.form("username")
site=request.form("site")
content=request.form("content")
if username<>"" and content<>"" then
objConn.execute("insert into dz_taotao (username,site,content) values ('"&username&"','"&site&"','"&content&"')")
response.write "添加成功<br><a href='admin_cmt.asp?a=list'>返回列表</a>"
response.end
else
response.write "<script>alert('能填的都要填啊！');history.back();</script>"
response.end
end if

case "p"
%>
<%
case "del"
t=request.QueryString("t")
id=request.QueryString("id")
page=request.QueryString("page")
if not isnumeric(id) and not isnumeric(t) then
response.write "<script>alert('操作提示：对不起，没有找到您要删除的信息！');history.back();</script>"
response.End()
end if
objConn.execute("delete from dz_comment where id="&id&"")

	'更新评论
	set cc=objConn.execute("select count(*) as c_count from dz_comment where tt_id="&t&" and itype=0")
	if not cc.eof then
	objConn.execute("update [dz_taotao] set comments = "&cc("c_count")&" where id = "&t&"")
	end if
	cc.close:set cc=nothing


response.Redirect("admin_cmt.asp?a=list&page="&page)
response.End()

case default
response.Redirect("admin_cmt.asp?a=list&page=1")
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

