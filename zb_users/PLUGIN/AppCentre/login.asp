<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!-- #include file="function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("AppCentre")=False Then Call ShowError(48)
Call LoadPluginXmlInfo("AppCentre")
Call AppCentre_InitConfig
If Request.QueryString("act")="save" Then

	Dim strSendTB,s

	Dim objPing
	Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP")

	objPing.open "POST",APPCENTRE_URL & "zb_users/plugin/appcentre_server/vaild.asp",False

	objPing.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objPing.SetRequestHeader "Cookie","username="&vbsescape(Request.Form("app_username"))&"; password="&vbsescape(MD5(Request.Form("app_password")))
	objPing.send ""

	s=objPing.responseText

	Set objPing = Nothing

	app_config.Write "DevelopUserName",Request.Form("app_username")
	app_config.Write "DevelopPassWord",s
	app_config.Save

	If s<>"" Then
		SetBlogHint_Custom("开发者您好,欢迎登陆到APP应用中心!")
		Response.Redirect "server.asp"
	Else
		SetBlogHint_Custom("用户名或密码输入错误!")
		Response.Redirect "login.asp"
	End If
ElseIf Request.QueryString("act")="logout" Then

	app_config.Write "DevelopUserName",""
	app_config.Write "DevelopPassWord",""
	app_config.Save
	Response.Redirect "login.asp"

End If
%>
<%
BlogTitle="应用中心-登录"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader">应用中心</div>
          <div class="SubMenu">
            <%AppCentre_SubMenu(3)%>
          </div>
          <div id="divMain2" style="margin:auto 0;">
            <form action="?act=save" method="post">
              <table width="60%" border="0">
                <tr height="32">
                  <td colspan="2" align="center">开发者请在这里填写您在"APP应用中心"的用户名和密码,并点登陆。</td>
                </tr>
                <tr height="32" align="center">
                  <td>用户名</td>
                  <td><input type="text" name="app_username" value="" style="width:90%"/></td>
                </tr>
                <tr height="32" align="center">
                  <td>密&nbsp;&nbsp;&nbsp;&nbsp;码</td>
                  <td><input type="password" name="app_password" value="" style="width:90%" /></td>
                </tr>
                <tr height="32" align="center">
                  <td colspan="2" align="center"><input type="submit" value="提交" class="button" /></td>
                </tr>
              </table>

            </form>
          </div>
        </div>
<%
	If login_pw<>"" Then
		Response.Write "<script type='text/javascript'>$('div.SubMenu a[href=\'login.asp\']').hide();</script>"
	End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->