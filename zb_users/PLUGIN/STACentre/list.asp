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
BlogTitle="静态中心配置插件"

Dim rewrite
rewrite=False

If ZC_STATIC_MODE="REWRITE" Then rewrite=True

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style>
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
  <div class="SubMenu"> <a href="main.asp"><span class="m-left">文章,页面设置</span></a><a href="list.asp"><span class="m-left m-now">分类,作者,TAGS,时间设置</span></a>
  </div>
  <div id="divMain2">
    <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
<form id="form" name="form" method="post" action="save.asp">
<input type="hidden" name="edtZC_STATIC_MODE" id="edtZC_STATIC_MODE" value="<%=ZC_STATIC_MODE%>" />
<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>
<tr><td><p  align='left'><b>·开启全静态化支持</b><br/><span class='note'>主机安装有ISAPI Rewrite或是URL Rewrite可以开启此选项</span></p></td><td><p><input id="rewrite" name="rewrite" style="" type="text" value="<%=rewrite%>" class="checkbox"/></p></td></tr>


<tr><td width='30%'><p align='left'><b>·首页的URL配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_DEFAULT_REGEX' name='edtZC_DEFAULT_REGEX' style='width:500px;' type='text' value='<%=ZC_DEFAULT_REGEX%>' /></p></td></tr>
<tr><td width='30%'><p align='left'><b>·分类页的URL配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_CATEGORY_REGEX' name='edtZC_CATEGORY_REGEX' style='width:500px;' type='text' value='<%=ZC_CATEGORY_REGEX%>' /></label></p></td></tr>
<tr><td width='30%'><p align='left'><b>·作者页的URL配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_USER_REGEX' name='edtZC_USER_REGEX' style='width:500px;' type='text' value='<%=ZC_USER_REGEX%>' /></label></p></td></tr>
<tr><td width='30%'><p align='left'><b>·TAGS页的URL配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_TAGS_REGEX' name='edtZC_TAGS_REGEX' style='width:500px;' type='text' value='<%=ZC_TAGS_REGEX%>' /></label></p></td></tr>
<tr><td width='30%'><p align='left'><b>·日期页的URL配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_DATE_REGEX' name='edtZC_DATE_REGEX' style='width:500px;' type='text' value='<%=ZC_DATE_REGEX%>' /></label></p></td></tr>
</table>
    <script type="text/javascript">
$(document).ready(function(){ 
	$("span.imgcheck").live("click",function(){
	
		if($(this).prev('input').val()=="True"){
			$(".content-box").show();
			$("#edtZC_DEFAULT_REGEX").val("{%host%}/default.html")
			$("#edtZC_CATEGORY_REGEX").val("{%host%}/category-{%id%}.html")
			$("#edtZC_USER_REGEX").val("{%host%}/author-{%id%}.html")
			$("#edtZC_TAGS_REGEX").val("{%host%}/tags-{%id%}.html")
			$("#edtZC_DATE_REGEX").val("{%host%}/{%date%}.html")
			$("#edtZC_STATIC_MODE").val("REWRITE")
		}else{
			$(".content-box").hide();
			$("#edtZC_DEFAULT_REGEX").val("{%host%}/catalog.asp")
			$("#edtZC_CATEGORY_REGEX").val("{%host%}/catalog.asp?cate={%id%}")
			$("#edtZC_USER_REGEX").val("{%host%}/catalog.asp?user={%id%}")
			$("#edtZC_TAGS_REGEX").val("{%host%}/catalog.asp?tags={%alias%}")
			$("#edtZC_DATE_REGEX").val("{%host%}/catalog.asp?date={%date%}")
			$("#edtZC_STATIC_MODE").val("ACTIVE")
		}
	});
})
	</script>

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
<p><span class="star">请在网站根目录创建httpd.ini文件并把相关内容复制进去.</span></p>
</div>


<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab2">
<pre>
<%=LoadFromFile(BlogPath &"\zb_users\PLUGIN\STACentre\iis6_3.html","utf-8")%>
</pre>
<p><span class="star">请在网站根目录创建.htaccess文件并把相关内容复制进去.</span></p>
</div>

<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab3">
<pre>
<%=TransferHTML(LoadFromFile(BlogPath &"\zb_users\PLUGIN\STACentre\iis7.html","utf-8"),"[html-format]")%>
</pre>
<p><span class="star">请在网站根目录创建web.config文件并把相关内容复制进去.</span></p>
</div>

				</div> <!-- End .content-box-content -->
				
			</div> <!-- End .content-box -->

<br/>
<input name="" type="submit" class="button" value="保存"/>
</form>



</div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
