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
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
	<%Call SubMenu(1)%>
  </div>
  <div id="divMain2">
   <script type="text/javascript">ActiveLeftMenu("aAppcentre");</script>
   <%
   Select Case Request.QueryString("act")
		Case "list","":ExportPluginList
'		Case "install":ExportPluginInstall
		Case "detail":ExportPluginDetail
   End Select
   %>
  </div>
</div>
<SCRIPT type="text/javascript">
function showhidediv(id){
	try{
		if(document.getElementById(id)){
		if(document.getElementById(id).style.display=='none'){
			document.getElementById(id).style.display='block';
		}else{
			document.getElementById(id).style.display='none';
		}
		}
	}catch(e){}
} 
</SCRIPT>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->



