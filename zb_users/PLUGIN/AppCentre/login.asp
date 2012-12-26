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
Call InitConfig
If Request.QueryString("act")="save" Then
	app_config.Write "DevelopUserName",Request.Form("app_username")
	app_config.Write "DevelopPassWord",Request.Form("app_password")
	app_config.Save
	Call SetBlogHint(True,Empty,Empty)
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
            <%SubMenu(3)%>
          </div>
          <div id="divMain2">
            <form action="?act=save" method="post" enctype="application/x-www-form-urlencoded">
            登陆成功后，没有任何提示，您就可以不再输入那些烦人的信息直接评论，并且您的评论将被打上特殊标记。
              <table width="100%" border="0">
                <tr height="32">
                  <td colspan="2" align="center">请在这里填写后台得到的信息。如果您没有开发者账号，请不要填写。</td>
                </tr>
                <tr height="32" align="center">
                  <td>用户名</td>
                  <td><input type="text" name="app_username" value="<%=app_config.Read("DevelopUserName")%>" style="width:100%"/></td>
                </tr>
                <tr height="32" align="center">
                  <td>MD5</td>
                  <td><input type="password" name="app_password" value="HaveSomeThing" style="width:100%" /></td>
                </tr>
                <tr height="32" align="center">
                  <td colspan="2" align="center"><input type="submit" value="提交" class="button"/></td>
                </tr>
              </table>
            </form>
          </div>
        </div>
        <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script> 
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->