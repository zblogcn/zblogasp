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
Dim intHighlight

Call System_Initialize()
Call CheckReference("")
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("AppCentre")=False Then Call ShowError(48)
Call LoadPluginXmlInfo("AppCentre")
Call AppCentre_InitConfig
BlogTitle="应用中心"

intHighlight=0


%>

<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain">
  <div id="ShowBlogHint">
	<%Call GetBlogHint()%>
  </div>
  <div class="divHeader">应用中心</div>
  <div class="SubMenu">
<%
If Request.QueryString("method")="check" Then
	intHighlight=2
Else
	intHighlight=0
End If

AppCentre_SubMenu(intHighlight)
%>
</div>
<div id="divMain2">
<%
If Request.QueryString("method")="" Then
	Call Server_Open("view")
Else
	Call Server_Open(Request.QueryString("method"))
End If
%>

  </div>
</div>
<script type="text/javascript">ActiveLeftMenu("aAppcentre");</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
