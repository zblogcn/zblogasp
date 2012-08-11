'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:   大猪(myllop)
'// 版权所有:    www.izhu.org
'// 技术支持:    myllop@qq.com
'// 程序名称:    大猪滔滔
'// 程序版本:    1.0
'///////////////////////////////////////////////////////////////////////////////
<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<%
dim server_v1,server_v2
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
response.write "error"
response.end
end if

%>
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

Call dztaotao_Initialize

dim old_tread,old_top
dim c,u,s,tid,e,img,s_img,cc
dim t,DZ_Rs,i,ss,strTaotao
dim add_rs,newID
dim isclass,cmt_count
t=replace(request("t"),"'","")
select case t
case "p"'提交淘淘信息
	c=request("c")
	u=request("u")
	s=request("s")
	img=request("img")
	s_img=request("s_img")
	if c="" or u="" then
	response.write "0"'信息没写全
	response.end
	end if
	'objConn.execute("insert into dz_taotao (username,site,content,img,s_img) values ('"&u&"','"&s&"','"&c&"','"&img&"','"&s_img&"')")
	
	set add_rs= Server.CreateObject("adodb.recordset")
	add_rs.open "select * from dz_taotao",objConn,1,3
	add_rs.addnew
	add_rs("username") = u
	add_rs("site") = s
	add_rs("content") = c
	add_rs("img") = img
	add_rs("s_img") = s_img
	add_rs("itype") = DZTAOTAO_CHK_VALUE
	add_rs.update
	newID = add_rs("id")
	add_rs.close:	set add_rs=nothing
	
	
	'读取最近10条淘淘信息
	Set DZ_Rs=objConn.Execute("SELECT top 10 * FROM [dz_taotao] WHERE [id]>0 ORDER BY [id] DESC")
	If (Not DZ_Rs.bof) And (Not DZ_Rs.eof) Then
	'strTaotao = strTaotao & "<link rel=""stylesheet"" href=""" & ZC_BLOG_HOST & "PLUGIN/dztaotao/images/i.css"" type=""text/css"" media=""screen"" />" & vbCrLf
		For i=1 to 10
			ss=TransferHTML(UBBCode(DZ_Rs("content"),"[face][link][autolink][font][code][image][media][flash]"),"[nohtml][vbCrlf][upload]")
			ss=Replace(ss,vbCrlf,"")
			if i>1 then isclass = "display:none"
	strTaotao=strTaotao & "<li><div class=""n_cmt_content""> <span class=""n_cmt_auth""><a href="""& DZ_Rs("site") & """>" & DZ_Rs("username") & "</a></span>  "&ss&"  <font class=""n_cmt_time"">"&DZ_Rs("addtime")&"</font><div style=""clear:both;""></div></div></li>"&vbCrlf
			

			DZ_Rs.MoveNext
			If DZ_Rs.eof Then Exit For
		Next
	End If
	DZ_Rs.close
	Set DZ_Rs=Nothing

	strTaotao=TransferHTML(strTaotao,"[no-asp]")

	
	Call SaveToFile(BlogPath & "/zb_users/include/dztaotao.asp",strTaotao,"utf-8",True)
	
	Call ClearGlobeCache()
	Call LoadGlobeCache()

	response.write newID'添加成功

case "r"'提交评论
	t=request("tid")
	c=request("c")
	u=request("u")
	s=request("s")
	e=request("e")

	if c ="" or t ="" or not isnumeric(t) then
	response.write "0"'信息没写全
	response.end
	end if
	if DZTAOTAO_CMTLIMIT_VALUE =1 then
		if request.Cookies("is_cmt"&t&"") <> "" then
		response.write "-111"'已经评论过
		response.End
		end if
	end if
	
	'response.write "ID值："&t&"<br>"&c&"<br>"&u&"<br>"&s&"<br>"&e
	'response.end
	'objConn.Execute("INSERT INTO [dz_comment]([tt_id],[u_sername],[u_site],[content]) VALUES ("&t&",'"&t_rndName&"','"&s&"','"&c&"')")
	
	set add_rs= Server.CreateObject("adodb.recordset")
	add_rs.open "select * from dz_comment",objConn,1,3
	add_rs.addnew
	add_rs("tt_id") = t
	add_rs("u_sername") = u
	add_rs("u_site") = s
	add_rs("content") = c
	add_rs("itype") = DZTAOTAO_CMTCHK_VALUE
	add_rs.update
	newID = add_rs("id")
	add_rs.close:	set add_rs=nothing
	
	'更新评论
	set cc=objConn.execute("select count(*) as c_count from dz_comment where tt_id="&t&" and itype=0")
	if not cc.eof then
	objConn.execute("update [dz_taotao] set comments = "&cc("c_count")&" where id = "&t&"")
	end if
	cc.close:set cc=nothing
	response.Cookies("is_cmt"&t&"") = t
	response.write newID'添加成功

case "dingup"'顶
	t=request("tid")
	if t ="" or not isnumeric(t) then
	response.write "0"'顶失败
	response.end
	end if
	if request.Cookies("ding_"&t&"") <> "" then
	response.write "0"'已经顶过
	response.End
	end if
	objConn.Execute("update dz_taotao set ttop=ttop+1 where id="&t&"")
	old_top = objConn.execute("select ttop from dz_taotao where id="&t&"",0,1)(0)
	response.Cookies("ding_"&t&"") = t
	response.write old_top'顶成功
	
case "dingdown"'踩
	t=request("tid")
	if t ="" or not isnumeric(t) then
	response.write "0"'踩失败
	response.end
	end if
	if request.Cookies("ding_"&t&"") <> "" then
	response.write "0"'已经踩过
	response.End
	end if
	objConn.Execute("update dz_taotao set tread=tread+1 where id="&t&"")
	old_tread = objConn.execute("select tread from dz_taotao where id="&t&"",0,1)(0)
	response.Cookies("ding_"&t&"") = t
	response.write old_tread'踩成功
end select






%>
