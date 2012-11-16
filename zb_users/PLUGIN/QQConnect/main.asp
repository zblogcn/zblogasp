<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
'
Call System_Initialize()
init_qqconnect()

If CheckPluginState("QQConnect")=False Then Call ShowError(48)
BlogTitle="QQ连接"

Select Case Request.QueryString("act")
	Case "callback"
		Select Case Request.QueryString("type")
			Case "connect"
					Call qqconnect.c.GetOpenId(qqconnect.c.CallBack)
					If BlogUser.Level=1 Then
						Call qqconnect.tconfig.write("Connect_OpenID",qqconnect.config.qqconnect.openid)
						Call qqconnect.tconfig.write("Connect_AccessToken",qqconnect.config.qqconnect.accesstoken)
						Call qqconnect.tconfig.Save
					End If
					SetBlogHint True,Empty,Empty
					Response.Redirect "main.asp"
			Case "weibo"
				If BlogUser.Level=1 Then
					Call qqconnect.t.Run(11,"","","","")
					Call qqconnect.tconfig.write("weibo_token",qqconnect.config.weibo.token)
					Call qqconnect.tconfig.write("weibo_secret",qqconnect.config.weibo.secret)
					Call qqconnect.tconfig.Save
					SetBlogHint True,Empty,Empty
					Response.Redirect "main.asp"
				End If
		End Select
		Response.End
	Case "logout"
		Select Case Request.QueryString("type")
			Case "connect"
				If BlogUser.Level=1 Then
					Call qqconnect.tconfig.write("Connect_OpenID","")
					Call qqconnect.tconfig.write("Connect_AccessToken","")
					Call qqconnect.tconfig.Save
				End If
				SetBlogHint True,Empty,Empty
				Response.Redirect "main.asp"
			Case "weibo"
				If BlogUser.Level=1 Then
					Call qqconnect.tconfig.write("weibo_token","")
					Call qqconnect.tconfig.write("weibo_secret","")
					Call qqconnect.tconfig.Save
					SetBlogHint True,Empty,Empty
					Response.Redirect "main.asp"
				End If
		End Select
End Select

Call CheckReference("")
%>
    
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
<div class="divHeader">QQ连接</div>
<div class="SubMenu"><%=qqconnect_navbar(0)%></div>
<div id="divMain2">
<table width="100%" border="1">
  <tr height="32">
    <td>
    
   <%
Dim tmpObject
If BlogUser.Level=1 Then
	If qqconnect.config.weibo.token="" Then
		Response.Write "<a href='" & qqconnect.t.Run(1,"","","","") & "'><img src='resources/wb_170_32.png'/></a>"
	Else
		Set tmpObject=qqconnect_json.toobject(qqconnect.t.api("http://open.t.qq.com/api/user/info","{}","GET"))
		Response.Write "欢迎回来，腾讯微博用户" & tmpObject.data.nick & "(" & tmpObject.data.name & ") <a href='main.asp?act=logout&type=weibo'>点击这里注销</a>"
	End If
End If


%></td>
  </tr>
  <tr height="32">
    <td>
<%
If qqconnect.config.qqconnect.appid<>"" Then
	If BlogUser.Level=1 Then
		qqconnect.config.qqconnect.openid=qqconnect.config.qqconnect.admin.openid
		qqconnect.config.qqconnect.accesstoken=qqconnect.config.qqconnect.admin.accesstoken
		If qqconnect.config.qqconnect.openid="" Then
			Response.Write "<a href='" & qqconnect.c.Authorize() & "'><img src='resources/logo_170_32.png'/></a>"
		Else
			Set tmpObject=qqconnect_json.toobject(qqconnect.c.api("https://graph.qq.com/user/get_user_info","{}","GET"))
			Response.Write "欢迎回来，QQ空间用户" & tmpObject.nickname & "<a href='main.asp?act=logout&type=connect'>点击这里注销</a>"
		End If
	Else
		'xxxxx
	End If
Else
	Response.Write "未配置QQ互联APPID，无法使用本功能。"
End If
%>    
</td>
  </tr>
</table>



</div>
</div>

<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
