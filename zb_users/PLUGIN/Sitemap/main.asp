<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
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
<%
Call System_Initialize()
Call SiteMap_Initialize
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("SiteMap")=False Then Call ShowError(48)
BlogTitle="SiteMap"
If Request.QueryString("act")="save" Then
	SiteMapc.Write "c",CStr(Request.Form("int"))
	SiteMapc.Save
End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> <a href="main.asp"><span class="m-left m-now">SiteMap</span></a>
  </div>
  <div id="divMain2">
    <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
<form id="form1" name="form1" method="post" action="?act=save">
  <label for="int">请输入文章数量，0为不限</label>
  <input type="number" name="int" id="int" style="width:100%" value="<%=SiteMapc.Read("c")%>" min="0"/><br/>
  <input name="a" type="submit" class="button" value="保存"/>
</form>

</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
