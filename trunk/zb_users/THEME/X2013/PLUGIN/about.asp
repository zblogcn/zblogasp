<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../../c_option.asp" -->
<!-- #include file="../../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../../plugin/p_config.asp" -->
<!-- #include file="Function.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("X2013")=False Then Call ShowError(48)
BlogTitle="X2013主题设置"
%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<style>
p{line-height:1.5em;padding:0.5em 0;}
</style>
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">
	<div id="ShowBlogHint"><%Call GetBlogHint()%></div>
	<!--<div class="divHeader"><%=BlogTitle%></div>-->
  	<div class="SubMenu"><%=X2013_SubMenu(2)%></div>
	<div id="divMain2">
	<script type="text/javascript">ActiveTopMenu("aX2013");</script> 
	<!--SetCon Star.-->
    <table width="100%" style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
	  <tr>
		<th width="20%">
	<ol>
  <li>1、本主题移植自WordPress同名主题<a href="http://lisizhang.com/x2013" target="_blank">X2013</a>,主题涉及的图片、CSS等版权归原作者<a href="http://lisizhang.com/" target="_blank">菠萝</a>所有，同时感谢好友菠萝为本主题所做的设计。</li>
  <li>2、本主题Z-Blog版本版权归<a href="http://imzhou.com"  target="_blank">未寒</a>所有，包括但不限于主题附带插件等。</li>
  <li>3、本主题自带一个设置插件，可设置博客前台部分显示资料，如果设置内容为空则不会显示该项内容。</li>
  <li>4、当切换为其他主题，即禁用本主题时，使用本主题时产生的所有配置文件将自动删除。</li>
</ol>
		
		</th>
	  </tr>
	</table>
	</div>
</div>
<!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->