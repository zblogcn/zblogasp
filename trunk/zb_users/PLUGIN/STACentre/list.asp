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
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("STACentre")=False Then Call ShowError(48)
BlogTitle="静态管理中心"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
pre{
	border:1px solid #ededed;
	margin:0px;
}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> <a href="main.asp"><span class="m-left">配置页面</span></a><a href="list.asp"><span class="m-left m-now">ReWrite规则</span></a>
  </div>
  <div id="divMain2">
<%If ZC_POST_STATIC_MODE="REWRITE" Or ZC_STATIC_MODE="REWRITE" Then%>
			<div class="content-box" <%=IIF(rewrite,"style='display:block;'","style='display:none;'")%>><!-- Start Content Box -->
				
				<div class="content-box-header">
			
					<ul class="content-box-tabs">

	<li><a href="#tab1" class="default-tab"><span>IIS6+ISAPI Rewrite 2.X</span></a></li>
	<li><a href="#tab2"><span>IIS6+ISAPI Rewrite 3.X</span></a></li>
	<li><a href="#tab3"><span>IIS7,7.5+Url Rewrite</span></a></li>
					</ul>
					
					<div class="clear"></div>
					
				</div> <!-- End .content-box-header -->

				<div class="content-box-content">
<div class="tab-content default-tab" style='border:none;padding:0px;margin:0;' id="tab1">
<pre>
<%=LoadFromFile(BlogPath &"\zb_users\PLUGIN\STACentre\iis6_2.html","utf-8")%>
</pre>
<hr/>
<p><span class="star">请在网站根目录创建httpd.ini文件并把相关内容复制进去.</span></p>
</div>


<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab2">
<pre>
<%=LoadFromFile(BlogPath &"\zb_users\PLUGIN\STACentre\iis6_3.html","utf-8")%>
</pre>
<hr/>
<p><span class="star">请在网站根目录创建.htaccess文件并把相关内容复制进去.</span></p>
</div>

<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab3">
<pre>
<%=TransferHTML(LoadFromFile(BlogPath &"\zb_users\PLUGIN\STACentre\iis7.html","utf-8"),"[html-format]")%>
</pre>
<hr/>
<p><span class="star">请在网站根目录创建web.config文件并把相关内容复制进去.</span></p>
</div>

				</div> <!-- End .content-box-content -->
				
			</div> <!-- End .content-box -->
<%Else%>
<hr/>
<p><b>文章及页面和分类页都没有启用动态模式+Rewrite支持,所以没有可用规则.</b></p>

<%End If%>

</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
