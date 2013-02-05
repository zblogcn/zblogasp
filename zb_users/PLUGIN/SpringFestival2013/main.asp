<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("SpringFestival2013")=False Then Call ShowError(48)
BlogTitle="迎春"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=SpringFestival2013_SubMenu(0)%></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("aPlugInMng");</script> 
            在这里写入后台管理页面代码
          </div>
        </div>
        <img id="toolTip" src="mouse.gif" style="display:block"/>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<script>
$("html").mousemove(function(event){
		$("#toolTip").css({ top: event.clientY-20, left: event.clientX-20,position:'absolute'});
});
document.getElementById("toolTip").oncontextmenu=function(){return false}
$("#toolTip").mousedown(function(){return false}).dblclick(function(){$(this).remove()})
</script>
<%Call System_Terminate()%>
