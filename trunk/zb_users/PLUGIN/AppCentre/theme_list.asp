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
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("AppCentre")=False Then Call ShowError(48)

BlogTitle="应用中心"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<link rel="stylesheet" href="images/style.css" type="text/css" media="screen" />
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
	<%Call SubMenu(1)%>
  </div>
  <div id="divMain2">
   <script type="text/javascript">ActiveLeftMenu("aAppcentre");$("#leftmenu #nav_appcentre.on span").css("background-image","url('<%=GetCurrentHost%>zb_users/plugin/appcentre/images/web2.png')")</script>
   <%
Response.Flush

Response.Write ListTheme(GetHTTPPage(APPCENTRE_URL & "?findapp=1&page=1&count=10"))
   %>
  </div>
</div>
<script type="text/javascript">

</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->