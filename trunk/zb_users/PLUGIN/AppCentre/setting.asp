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

	enable_develop=Request.Form("app_enabledevelop")
	disable_check=Request.Form("app_disablecheck")
	app_config.Write "EnableDevelop",enable_develop
	app_config.Write "DisableCheck",disable_check
	app_config.Save

	Call SetBlogHint_Custom("设置成功.")

ElseIf Request.QueryString("act")="login" Then

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
		Response.Redirect "setting.asp"
	End If
ElseIf Request.QueryString("act")="logout" Then

	app_config.Write "DevelopUserName",""
	app_config.Write "DevelopPassWord",""
	app_config.Save
	Response.Redirect "setting.asp"

End If
%>
<%
BlogTitle="应用中心-设置与开发者登录"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu">
            <%AppCentre_SubMenu(1)%>
          </div>

          <div id="divMain2" style="margin:auto 0;">

            <form action="?act=save" method="post">
              <table width="100%" border="0">
                <tr height="32">
                  <th colspan="2" align="center">设置</td>
                </tr>
                <tr height="32">
                  <td width="60%" align="left"><b>启用开发者模式</b><br/><small>(启用开发者模式可以修改,导出应用,注册开发者还可以远程提交应用到APP应用中心)</small></td>
                  <td><input id="app_enabledevelop" name="app_enabledevelop" style="" type="text" value="<%=enable_develop%>" class="checkbox"/></td>
                </tr>
                <tr height="32">
                  <td width="60%" align="left"><b>禁用自动检查更新</b><br/><small>(禁用自动检查后,需要手动检查应用更新和系统更新)</small></td>
                  <td><input id="app_disablecheck" name="app_disablecheck" style="" type="text" value="<%=disable_check%>" class="checkbox"/></td>
                </tr>

              </table>
<hr/>
<p><input type="submit" value="提交" class="button" /></p>
<hr/>
            </form>

<div class="divHeader2">开发者登录</div>


<%
	If login_pw<>"" Then
%>
            <form action="?act=logout" method="post">
<p>开发者 <b><%=login_un%></b> 您好,您已经在当前客户端登录Z-Blog官方网站-APP应用中心.</p>
<p><input type="submit" value="退出登录" class="button" /></p>

            </form>
<%
	Else
%>
            <form action="?act=login" method="post">
              <table style="line-height:3em;" width="100%" border="0">
                <tr height="32">
                  <th  align="center">开发者请填写您在"APP应用中心"的用户名和密码并点登陆</td>
                </tr>
                <tr height="32">
                  <td  align="center">用户名:<input type="text" name="app_username" value="" style="width:40%"/></td>
                </tr>
                <tr height="32">
                  <td  align="center">密&nbsp;&nbsp;&nbsp;&nbsp;码:<input type="password" name="app_password" value="" style="width:40%" /></td>
                </tr>
                <tr height="32" align="center">
                  <td align="center"><input type="submit" value="登陆" class="button" /></td>
                </tr>
              </table>
            </form>
<%
	End If
%>


          </div>
        </div>

<script type="text/javascript">ActiveLeftMenu("aAppcentre");</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->