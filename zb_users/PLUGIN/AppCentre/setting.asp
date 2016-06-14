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
	check_beta=Request.Form("app_checkbeta")
	app_config.Write "EnableDevelop",enable_develop
	app_config.Write "DisableCheck",disable_check
	app_config.Write "CheckBeta",check_beta
	app_config.Save

	Call SetBlogHint_Custom("设置成功.")

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
                  <th colspan="2" align="center">设置
                    </td>
                </tr>
                <tr height="32">
                  <td width="30%" align="left"><p><b>· 启用开发者模式</b><br/>
                      <span class="note">&nbsp;&nbsp;启用开发者模式可以修改应用信息、导出应用和远程提交应用</span></p></td>
                  <td><input id="app_enabledevelop" name="app_enabledevelop" type="text" value="<%=enable_develop%>" class="checkbox"/></td>
                </tr>
                <tr height="32">
                  <td width="30%" align="left"><p><b>· 禁用自动检查更新</b><br/>
                      <span class="note">&nbsp;&nbsp;禁用自动检查后,需要手动检查应用更新和系统更新 </span></p></td>
                  <td><input id="app_disablecheck" name="app_disablecheck" type="text" value="<%=disable_check%>" class="checkbox"/></td>
                </tr>
                <tr height="32">
                  <td width="30%" align="left"><p><b>· 检查Beta版程序</b><br/>
                      <span class="note">&nbsp;&nbsp;若打开，则系统将检查最新测试版的Z-Blog更新</span></p></td>
                  <td><input id="app_checkbeta" name="app_checkbeta" type="text" value="<%=check_beta%>" class="checkbox"/></td>
                </tr>
              </table>
              <hr/>
              <p>
                <input type="submit" value="提交" class="button" />
              </p>
              <hr/>
            </form>
        </div>
        <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script> 
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->