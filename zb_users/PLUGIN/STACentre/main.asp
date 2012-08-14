<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
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
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("STACentre")=False Then Call ShowError(48)
BlogTitle="静态中心配置插件"
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> <a href="main.asp"><span class="m-left m-now">文章,页面设置</span></a><a href="list.asp"><span class="m-left">分类,作者,TAGS,时间设置</span></a>
  </div>
  <div id="divMain2">
    <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
<form id="form" name="form" method="post" action="save.asp">

<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>
<tr><td width='30%'><p align='left'><b>·文章类的静态配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_ARTICLE_REGEXT' name='edtZC_ARTICLE_REGEX' style='width:500px;' type='text' value='<%=ZC_ARTICLE_REGEX%>' /></p></td></tr>
<tr><td width='30%'><p>推荐配置</p></td><td>
<p><label onclick="changeval(1,1)"><input type="radio" name="radio" />&nbsp;&nbsp;配置1:文章名型(默认) http://www.yourblog/post/articlename.html</label></p>
<p><label onclick="changeval(1,2)"><input type="radio" name="radio" />&nbsp;&nbsp;配置2:日期+文章名型 http://www.yourblog/2012/08/articlename.html</label></p>
<p><label onclick="changeval(1,3)"><input type="radio" name="radio" />&nbsp;&nbsp;配置3:分类别名+文章名型 http://www.yourblog/categroyname/articlename.html</label></p>
<p><label onclick="changeval(1,4)"><input type="radio" name="radio" />&nbsp;&nbsp;配置4:文章名目录型 http://www.yourblog/post/articlename/</label></p>
<p><label onclick="changeval(1,5)"><input type="radio" name="radio" />&nbsp;&nbsp;配置5:分类别名+文章ID目录型 http://www.yourblog/categroyname/123/</label></p>
</td></tr>
<tr><td width='30%'><p align='left'><b>·页面类的静态配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_PAGE_REGEX' name='edtZC_PAGE_REGEX' style='width:500px;' type='text' value='<%=ZC_PAGE_REGEX%>' /></label></p></td></tr>
<tr><td width='30%'><p>推荐配置</p></td><td>
<p><label onclick="changeval(2,1)"><input type="radio" name="radio2" />&nbsp;&nbsp;配置1:页面名型(默认) http://www.yourblog/pagename.html</label></p>
<p><label onclick="changeval(2,2)"><input type="radio" name="radio2" />&nbsp;&nbsp;配置2:页面名目录型 http://www.yourblog/pagename/</label></p>
</td></tr>
</table>
<p><span class="note">您可以自定义静态配置,配置必须是{%host%}打头,".html"结尾,中间可以用{%post%},{%category%},{%user%},{%year%},{%month%},{%day%},{%id%},{%alias%}组合.</span></p>
<p><span class="note">{%post%}为文章发布目录,可以在网站设置里修改,{%category%}为文章的别名,{%user%}是用户别名,{%alias%}是文章别名,没有设置文章别名系统会自动采用ID填充.</span></p>
<br/>
<input name="" type="submit" class="button" value="保存"/>
</form>
</div>
</div>
<script type="text/javascript">
function changeval(a,b){
	if(a==1){
		a="#edtZC_ARTICLE_REGEXT";
		if(b==1){b="{%host%}/{%post%}/{%alias%}.html"};
		if(b==2){b="{%host%}/{%year%}/{%month%}/{%alias%}.html"};
		if(b==3){b="{%host%}/{%category%}/{%alias%}.html"};
		if(b==4){b="{%host%}/{%post%}/{%alias%}/default.html"};
		if(b==5){b="{%host%}/{%category%}/{%id%}/default.html"};
	}else{
		a="#edtZC_PAGE_REGEX";
		if(b==1){b="{%host%}/{%alias%}.html"};
		if(b==2){b="{%host%}/{%alias%}/default.html"};
	}
	$(a).val(b);
}
</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
