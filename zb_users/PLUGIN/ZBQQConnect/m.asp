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
Call ZBQQConnect_Initialize()

Call CheckReference("")

If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ZBQQConnect")=False Then Call ShowError(48)


%>
    
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader">ZBQQConnect</div>
<div class="SubMenu" style="border-bottom:5px solid #3399cc;"><%=ZBQQConnect_SBar(2)%></div>
<div id="divMain2">

      <%
If Request.QueryString("act")="del" Then Call DelQQ(Request.QueryString("id")) Else Call ExportQQList(Request.QueryString("page"),Request("qq_id"),Request("qq_uid"),Request("qq_oid"),Request("qq_un"))
Function ExportQQList(intPage,intId,intUid,strOid,strUn)

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
	Call CheckParameter(strUn,"sql",-1)
	
	
	strOid=FilterSQL(strOid)
	strUn=FilterSQL(strUn)


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
		If strUn<>"-1" Then
			If ZC_MSSQL_Enable=False Then
				strSQL = strSQL & " AND  (InStr(1,[QQ_Eml],LCase('" & strUn &"'),0)<>0)"
			Else
				strSQL = strSQL & " AND  (CHARINDEX('" & strUn &"',[QQ_Eml])<>0)"
			End If
		End If
		Response.Write "<form class=""search"" id=""edit"" method=""post"" action=""m.asp"">"
		Response.Write "<p>搜索符合条件的用户</p><p>"
		Response.Write "ID<input type=""text"" name=""qq_id"" style=""width:30px;"" value="""&IIf(intId<>-1,intId,"")&""" />&nbsp;&nbsp;"
		Response.Write "用户ID<input type=""text"" name=""qq_uid"" style=""width:30px;"" value="""&IIf(intUid<>-1,intuId,"")&"""/>&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write "E-Mail<input id=""qq_un"" name=""qq_un"" style=""width:250px;"" type=""text"" value="""&IIf(strUn<>"-1",strUn,"")&""" /> "		
		Response.Write "OpenID<input type=""text"" name=""qq_oid"" style=""width:250px;"" value="""&IIf(strOid<>"-1",strOid,"")&""" />&nbsp;&nbsp;&nbsp;&nbsp;"
		Response.Write "<input type=""submit"" class=""button"" value="""&ZC_MSG087&"""/>"
		Response.Write "</p></form>"
	Else
		strSQL= strSQL & "WHERE ([QQ_UserID] = " & BlogUser.ID&")"
	End If


	Response.Write "<table border=""1"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""tableBorder"">"
	Response.Write "<tr><th width=""5%"">ID</th><th width=""14%"">ID|绑定用户</th><th width=""14%"">E-Mail</th><th width=""14%"">OpenID</th><th>AccessToken</th><th width=""14%"">空间头像</th><th width=""14%"">微博头像</th><th width=""14%"">Gravatar</th><th width=""14%""></th></tr>"
'response.write strsql
	objRS.Open("SELECT * FROM [blog_Plugin_ZBQQConnect] "& strSQL &" ORDER BY [QQ_ID] ASC")
	objRS.PageSize=ZC_MANAGE_COUNT
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount
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
					End If
				End If
			Next
			Response.Write "<td>" & objRs("QQ_Eml") & "</td>"

			Response.Write "<td>" & objRs("QQ_OpenID") & "</td>"
			Response.Write "<td>" & objRs("QQ_AToken") & "</td>"
			Response.Write "<td><img src=""" & objRs("QQ_QZoneHead") & """ width=""32px"" height=""32px""/></td>"
			Response.Write "<td><img src="""&objRs("QQ_THead")&"/30"" width=""32px"" height=""32px""/></td>"
			Response.Write "<td><img src=""http://www.gravatar.com/avatar/"&MD5(objRs("QQ_Eml"))&""" width=""32px"" height=""32px""/></td>"
			Response.Write "<td><a onclick='return window.confirm("""& ZC_MSG058 &""");' href=""m.asp?act=del&id=" & objRS("Qq_ID") & """><img src=""../../../ZB_SYSTEM/image/admin/delete.png"" alt=""" & ZC_MSG063 & """ title=""" & ZC_MSG063 & """ width=""16"" /></a></td>"
			Response.Write "</tr>"

			objRS.MoveNext
			If objRS.eof Then Exit For

		Next

	End If

	Response.Write "</table>"

	strPage=ExportPageBar(intPage,intPageAll,ZC_PAGEBAR_COUNT,"m.asp?qq_id="&intid&"&qq_uid="&intuid&"&stroid="&stroid&"&strUn="&strun&"&page=")

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

<%
'以下为example页面代码，对SDK开发无用
'导航栏生成 

'空转判断
function pdkz(text)
	if text=null or text=empty or text="" then pdkz="空转" else pdkz=text
end function
set ZBQQConnect_class=nothing
%>