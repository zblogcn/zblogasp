<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
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
Dim intHighlight

Call System_Initialize()
Call CheckReference("")
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("AppCentre")=False Then Call ShowError(48)
Call LoadPluginXmlInfo("AppCentre")
Call AppCentre_InitConfig

If shop_un="" Or shop_pw="" Then
	BlogTitle="应用中心-登录应用商城"
Else
	BlogTitle="应用中心-我的应用仓库"
End If

intHighlight=0


If Request.QueryString("act")="shoplogin" Then

	Dim s

	Call Server_Open("shopvaild")
	
	s=strResponse

	app_config.Write "ShopUserName",Request.Form("shop_username")
	app_config.Write "ShopPassWord",s
	app_config.Save

	
	If s<>"" Then
		SetBlogHint_Custom("您已成功登录""应用中心""商城.")
		Response.Redirect "client.asp"
	Else
		SetBlogHint_Custom("购买者账户名或密码错误.")
		Response.Redirect "client.asp"
	End If
ElseIf Request.QueryString("act")="shoplogout" Then

	app_config.Write "ShopUserName",""
	app_config.Write "ShopPassWord",""
	app_config.Save
	Response.Redirect "server.asp"

End If



%>

<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain">
  <div id="ShowBlogHint">
	<%Call GetBlogHint()%>
  </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu">
<%

intHighlight=6


AppCentre_SubMenu(intHighlight)
%>
</div>
<div id="divMain2">
<%
If app_config.read("ShopUserName")="" Or app_config.read("ShopPassWord")="" Then
%>
            <form action="?act=shoplogin" method="post">
              <table style="line-height:3em;" width="100%" border="0">
                <tr height="32">
                  <th  align="center">请填写您在"<a href="http://app.rainbowsoft.org/?shop&type=account" target="_blank">应用中心</a>"的购买者账号(Email)和密码
                    </th>
                </tr>
                <tr height="32">
                  <td  align="center">&nbsp;&nbsp;账号:
                    <input type="text" name="shop_username" value="" style="width:40%"/></td>
                </tr>
                <tr height="32">
                  <td  align="center">&nbsp;&nbsp;密码:
                    <input type="password" name="shop_password" value="" style="width:40%" /></td>
                </tr>
                <tr height="32" align="center">
                  <td align="center"><input type="submit" value="登陆" class="button" /></td>
                </tr>
              </table>
            </form>
<%
Else
	Call Server_Open("shoplist")
End If
%>

  </div>
</div>
<script type="text/javascript">ActiveLeftMenu("aAppcentre");</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
