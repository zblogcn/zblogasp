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


Dim ID

ID=Request.QueryString("id")

BlogTitle="应用中心-主题编辑"


%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
	<%Call SubMenu(3)%>
  </div>
  <div id="divMain2">

<table border="1" width="100%" cellspacing="0" cellpadding="0" class="tableBorder tableBorder-thcenter">
<tr><th width='20%'>&nbsp;</th><th>&nbsp;</th></tr>

<tr><td><b>· 主题ID</b></td><td>&nbsp;</td></tr>
<tr><td><b>· 主题名称</b></td><td>&nbsp;</td></tr>
<tr><td><b>· 主题发布页面</b></td><td>&nbsp;</td></tr>
<tr><td><b>· 主题简介</b></td><td>&nbsp;</td></tr>
<tr><td><b>· 适用的 Z-Blog 版本</b></td><td>&nbsp;</td></tr>
<tr><td><b>· 主题版本号</b></td><td>&nbsp;</td></tr>
<tr><td><b>· 主题首发时间</b></td><td>&nbsp;</td></tr>
<tr><td><b>· 主题最后修改时间</b></td><td>&nbsp;</td></tr>

<tr><td><b>· 作者名称</b></td><td>&nbsp;</td></tr>
<tr><td><b>· 作者邮箱</b></td><td>&nbsp;</td></tr>
<tr><td><b>· 作者网站</b></td><td>&nbsp;</td></tr>
<tr><td><b>· 源作者名称</b> (可选)</td><td>&nbsp;</td></tr>
<tr><td><b>· 源作者邮箱</b> (可选)</td><td>&nbsp;</td></tr>
<tr><td><b>· 源作者网站</b> (可选)</td><td>&nbsp;</td></tr>
<tr><td><b>· 详细说明</b> (可选)</td><td>&nbsp;</td></tr>

<tr><td><b>· 内置插件简介</b> (可选)</td><td>&nbsp;</td></tr>
<tr><td><b>· 内置插件管理页面</b> (可选)</td><td>&nbsp;</td></tr>
<tr><td><b>· 内置插件管理权限</b> (可选)</td><td>&nbsp;</td></tr>

</table>

  </div>
</div>
   <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->