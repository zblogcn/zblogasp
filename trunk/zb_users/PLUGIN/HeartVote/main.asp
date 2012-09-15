<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    Z-Blog 1.7
'// 插件制作:    
'// 备    注:    
'// 最后修改：   
'// 最后版本:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>4 Then Call ShowError(6) 

If CheckPluginState("HeartVote")=False Then Call ShowError(48)

BlogTitle="用“心”打分"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

			<div id="divMain">
<div id="ShowBlogHint">
      <%Call GetBlogHint()%>
</div>
<div class="divHeader2"><%=BlogTitle%></div>
<div id="divMain2">
<form id="edit" name="edit" method="post">
<p>使用说明请访问：</p>
<p><a href="http://bbs.rainbowsoft.org/thread-19200-1-1.html" target="_blank">http://bbs.rainbowsoft.org/thread-19200-1-1.html</a></p>
</form>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>

