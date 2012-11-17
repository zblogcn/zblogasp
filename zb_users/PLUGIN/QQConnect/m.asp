<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_manage.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
'
Call System_Initialize()
init_QQConnect()
Call CheckReference("")

If BlogUser.Level=5 Then Call ShowError(6)
If CheckPluginState("QQConnect")=False Then Call ShowError(48)

BlogTitle="QQ互联-绑定管理"
%>      <%
Sub DelQQ(id)
	QQConnect_DB.ID=id
	QQConnect_DB.Del
	Call SetBlogHint(True,True,Empty)
End Sub
If Request.QueryString("act")="del" Then Call DelQQ(Request.QueryString("id"))  %>
    
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader">QQ互联</div>
<div class="SubMenu"><%=qqconnect.functions.navbar(2)%></div>
<div id="divMain2">

<%
Call ExportQQList(Request.QueryString("page"),Request("qq_id"),Request("qq_uid"),Request("qq_oid"))
Function ExportQQList(intPage,intId,intUid,strOid)

'Call Add_Response_Plugin("Response_Plugin_ArticleMng_SubMenu",MakeSubMenu(ZC_MSG168 & "","../cmd.asp?act=ArticleEdt&amp;webedit=" & ZC_BLOG_WEBEDIT,"m-left",False))

	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim intPageAll

	Call CheckParameter(intPage,"int",1)
	Call CheckParameter(intId,"int",-1)
	Call CheckParameter(intUid,"int",-1)
	Call CheckParameter(strOid,"sql",-1)
	
	
	strOid=FilterSQL(strOid)


	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""
	If CheckRights("Root") Then
		If intId>-1 Then
			strSQL="WHERE ([QQ_ID]="&intId&")"
		Else
			strSQL="WHERE ([QQ_ID]>0)"
		End If
		If strOid<>"-1" Then
			If ZC_MSSQL_Enable=False Then
				strSQL = strSQL & " AND  (InStr(1,[QQ_OpenID],LCase('" & strOid &"'),0)<>0)"
			Else
				strSQL = strSQL &   " AND  (CHARINDEX('" & strOid &"',[QQ_OpenID])<>0)"
			End If
		End If
		If intUid>-1 Then
				strSQL=strSQL & " AND ([QQ_UserID]="&intUid&")"
		End If

		Response.Write "<form class=""search"" id=""edit"" method=""post"" action=""m.asp"">"
		Response.Write "<p>搜索符合条件的用户</p><p>"
		Response.Write "ID:&nbsp;<input type=""text"" name=""qq_id"" style=""width:30px;"" value="""&IIf(intId<>-1,intId,"")&""" />&nbsp;&nbsp;"
		Response.Write "用户ID:&nbsp;<input type=""text"" name=""qq_uid"" style=""width:30px;"" value="""&IIf(intUid<>-1,intuId,"")&"""/>&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write "OpenID:&nbsp;<input type=""text"" name=""qq_oid"" style=""width:220px;"" value="""&IIf(strOid<>"-1",strOid,"")&""" />&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write "<input type=""submit"" class=""button"" value="""&ZC_MSG087&"""/>"
		Response.Write "</p></form>"
	Else
		strSQL= strSQL & "WHERE ([QQ_UserID] = " & BlogUser.ID&")"
	End If


	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder"">"
	Response.Write "<tr><th width=""5%"">ID</th><th width=""30%"">ID|绑定用户</th><th width=""14%"">OpenID</th><th>AccessToken</th><th width=""14%"">空间头像</th><th width=""14%"">微博头像</th><th width=""14%"">Gravatar</th><th width=""14%""></th></tr>"
'response.write strsql
	objRS.Open("SELECT * FROM [blog_Plugin_QQConnect] "& strSQL &" ORDER BY [QQ_ID] ASC")
	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount
	Response.Write "<tr><td>0</td><td><b>[固定]</b>管理员("&BlogUser.FIrstName&")</td><td>"&qqconnect.config.qqconnect.admin.openid
	Response.Write "</td><td>"&qqconnect.config.qqconnect.admin.accesstoken
	Response.Write "</td><td>"
	Response.Write IIf(BlogUser.Meta.GetValue("QQConnect_Head1")="","","<img src="""&BlogUser.Meta.GetValue("QQConnect_Head1")&""" width=""32px"" height=""32px""/>")
	Response.Write "</td><td>"
	Response.Write IIf(BlogUser.Meta.GetValue("QQConnect_Head2")="","","<img src="""&BlogUser.Meta.GetValue("QQConnect_Head2")&"/30"" width=""32px"" height=""32px""/>")
	Response.Write "</td></td><td>"
	Response.Write "<img src=""http://www.gravatar.com/avatar/"& MD5(BlogUser.Email) & """ width=""32px"" height=""32px""/>"
	Response.Write "</td><td></td></tr>"
	If (Not objRS.bof) And (Not objRS.eof) Then
		For i=1 to objRS.PageSize
			Response.Write "<tr>"
			Response.Write "<td>" & objRS("QQ_ID") & "</td>"
			Call GetUser
			Dim User
			For Each User in Users
				If IsObject(User) Then
					If User.ID=objRS("QQ_UserID") Then
						Response.Write "<td>"&User.Name&"|" & User.ID & "</td>"
						Response.Write "<td>" & objRs("QQ_OpenID") & "</td>"
						Response.Write "<td>" & objRs("QQ_AToken") & "</td>"
						Response.Write "<td><img src=""" & User.Meta.GetValue("QQConnect_Head1") & """ width=""32px"" height=""32px""/></td>"
						Response.Write "<td><img src=""" & User.Meta.GetValue("QQConnect_Head2")&"/30"" width=""32px"" height=""32px""/></td>"
						Response.Write "<td><img src=""http://www.gravatar.com/avatar/"& MD5(BlogUser.Email) & """ width=""32px"" height=""32px""/></td>"
						Response.Write "<td><a href=""edit.asp?id="&objRs("QQ_ID")&"""><img src=""../../../ZB_SYSTEM/image/admin/page_edit.png"" title=""编辑""/></a>  <a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""m.asp?act=del&id=" & objRS("Qq_ID") & """><img src=""../../../ZB_SYSTEM/image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
						Response.Write "</tr>"
					End If
				End If
			Next
			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	Response.Write "</table>"

	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"m.asp?qq_id="&intid&"&qq_uid="&intuid&"&stroid="&stroid&"&page=")

	Response.Write "<hr/><p class=""pagebar"">" & ZC_MSG042 & ": " & strPage & "</p>"
	Response.Write "</div>"


	objRS.Close
	Set objRS=Nothing

	ExportQQList=True

End Function


%>


</div>
</div>

<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

