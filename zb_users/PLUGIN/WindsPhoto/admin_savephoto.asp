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

<!-- #include file="function.asp" -->

<%
Call System_Initialize
Call WindsPhoto_Initialize()
%><%
'检查非法链接
'Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

%>

<%
If Request.QueryString("action") = "" Then
    Call SetBlogHint_Custom("!! 参数错误.")
    Response.Redirect"admin_main.asp"
Else
    action = Request.QueryString("action")
    id = Request.QueryString("id")
	typeid = Request.QueryString("typeid")
	tt = Request.QueryString("t")
End If

If action = "hot" then

	If tt = "" Then tt = 0

	Set rs = Server.CreateObject("ADODB.RecordSet")
	sql = "select * FROM WindsPhoto_desktop where id="&id
	rs.Open sql, objConn, 1, 3
	If rs("hot") = "" Or IsNull(rs("hot")) = TRUE Then
		rs("hot") = "0"
	Else
		rs("hot") = tt
	End If

	rs.update
	rs.Close
	Set rs = Nothing

	Call SetBlogHint_Custom("? 设置封面成功，如果需要重新生成静态首页，<a href="""&ZC_BLOG_HOST&"zb_users/plugin/WindsPhoto/admin_html.asp"">请点击这里更新</a>。")

	Response.Redirect "admin_addphoto.asp?typeid=" & typeid

elseif action = "del" Then

    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * FROM WindsPhoto_desktop where id="&id
    rs.Open sql, objConn, 1, 3

    fn = rs("url")
    fn1 = rs("surl")

    If Left(fn, 4)<>"http" Then

        filepath = Server.MapPath(fn)
        Set Fso = server.CreateObject("scripting.filesystemobject")
        Fso.DeleteFile(Filepath)

        If fn1<>fn Then
            filepath = Server.MapPath(fn1)
            Set Fso = server.CreateObject("scripting.filesystemobject")
            Fso.DeleteFile(Filepath)
        End If

    End If
    objconn.Execute "delete FROM WindsPhoto_desktop where id="&id

	Call SetBlogHint_Custom("√ 删除照片成功.")

	Response.Redirect "admin_addphoto.asp?typeid=" & typeid

Else

    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * FROM WindsPhoto_desktop where id="&Request.QueryString("id")
    rs.Open sql, objConn, 1, 3
    rs("name") = Request.Form("name")
    rs("url") = Request.Form("url")
    rs("surl") = Request.Form("surl")
    rs("zhuanti") = Request.Form("zhuanti")
    rs("hot") = Request.Form("hot")
    rs("jj") = Request.Form("jj")

    if Request.Form("itime")<>"" then
      itime = Request.Form("itime")
    else
      itime =now()
    end if

    rs("itime") = itime
    rs.update
    rs.Close
    Set rs = Nothing

    Call SaveLastest()

    conn.Close
    Set conn = Nothing

    Call SetBlogHint_Custom("√ 编辑照片信息成功</a>")
    'Response.Redirect "admin_editphoto.asp?id=" & Request.QueryString("id") & "&action=edit"
    Response.Redirect "admin_addphoto.asp?typeid=" & Request.Form("zhuanti")
End If
%>

<%
Call System_Terminate()

'If Err.Number<>0 Then
    'Call ShowError(0)
'End If
%>