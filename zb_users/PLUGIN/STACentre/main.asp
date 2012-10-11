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

<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> <a href="main.asp"><span class="m-left m-now">配置页面</span></a><a href="list.asp"><span class="m-left">ReWrite规则</span></a>
  </div>
  <div id="divMain2">
    <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
<form id="form" name="form" method="post" action="save.asp">




			<div class="content-box"><!-- Start Content Box -->
				
				<div class="content-box-header">
			
					<ul class="content-box-tabs">

	<li><a href="#tab1" class="default-tab"><span>文章及页面静态化设置</span></a></li>
	<li><a href="#tab2"><span>分类页静态化设置</span></a></li>
					</ul>
					
					<div class="clear"></div>
					
				</div> <!-- End .content-box-header -->
				
				<div class="content-box-content">


<div class="tab-content default-tab" style='border:none;padding:0px;margin:0;' id="tab1">

<input type="hidden" name="edtZC_POST_STATIC_MODE" id="edtZC_POST_STATIC_MODE" value="<%=ZC_POST_STATIC_MODE%>" />
<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>
<tr><td width='30%'><p align='left'><b>·文章,页面的静态化选项</b><br/><span class='note'>主机安装有ISAPI Rewrite或是URL Rewrite可以开启Rewrite选项</span></p></td><td><p><label><input type="radio" value="STATIC" name="POST_STATIC" <%=IIF(ZC_POST_STATIC_MODE="STATIC","checked='checked'","")%>/>&nbsp;&nbsp;1.静态页面(系统默认)</label>&nbsp;&nbsp;&nbsp;&nbsp;<label><input type="radio" value="ACTIVE" name="POST_STATIC" <%=IIF(ZC_POST_STATIC_MODE="ACTIVE","checked='checked'","")%>/>&nbsp;&nbsp;2.动态页面</label>&nbsp;&nbsp;&nbsp;&nbsp;<label><input type="radio" value="REWRITE" name="POST_STATIC" <%=IIF(ZC_POST_STATIC_MODE="REWRITE","checked='checked'","")%>/>&nbsp;&nbsp;3.动态页面+Rewrite支持</label></p></td></tr>
<tr><td width='30%'><p align='left'><b>·文章类的静态配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_ARTICLE_REGEXT' name='edtZC_ARTICLE_REGEX' style='width:500px;' type='text' value='<%=ZC_ARTICLE_REGEX%>' /></p></td></tr>
<tr><td width='30%'><p>推荐配置</p></td><td>
<p><label onClick="changeval(1,1)"><input type="radio" name="radio" />&nbsp;&nbsp;配置1:文章名型(默认) http://www.yourblog/post/articlename.html</label></p>
<p><label onClick="changeval(1,2)"><input type="radio" name="radio" />&nbsp;&nbsp;配置2:日期+文章名型 http://www.yourblog/2012/08/articlename.html</label></p>
<p><label onClick="changeval(1,3)"><input type="radio" name="radio" />&nbsp;&nbsp;配置3:分类别名+文章名型 http://www.yourblog/categroyname/articlename.html</label></p>
<p><label onClick="changeval(1,4)"><input type="radio" name="radio" />&nbsp;&nbsp;配置4:文章名目录型 http://www.yourblog/post/articlename/</label></p>
<p><label onClick="changeval(1,5)"><input type="radio" name="radio" />&nbsp;&nbsp;配置5:分类别名+文章ID目录型 http://www.yourblog/categroyname/123/</label></p>
</td></tr>
<tr><td width='30%'><p align='left'><b>·页面类的静态配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_PAGE_REGEX' name='edtZC_PAGE_REGEX' style='width:500px;' type='text' value='<%=ZC_PAGE_REGEX%>' /></label></p></td></tr>
<tr><td width='30%'><p>推荐配置</p></td><td>
<p><label onClick="changeval(2,1)"><input type="radio" name="radio2" />&nbsp;&nbsp;配置1:页面名型(默认) http://www.yourblog/pagename.html</label></p>
<p><label onClick="changeval(2,2)"><input type="radio" name="radio2" />&nbsp;&nbsp;配置2:页面名目录型 http://www.yourblog/pagename/</label></p>
</td></tr>
</table>
<p><span class="note">您可以自定义静态配置,配置必须是{%host%}打头,".html"结尾,中间可以用{%post%},{%category%},{%author%},{%year%},{%month%},{%day%},{%id%},{%alias%}组合.</span></p>
<p><span class="note">{%post%}为文章发布目录,可以在网站设置里修改,{%category%}为文章的别名,{%user%}是用户别名,{%alias%}是文章别名,没有设置文章别名系统会自动采用ID填充.</span></p>

</div>

<div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab2">

<input type="hidden" name="edtZC_STATIC_MODE" id="edtZC_STATIC_MODE" value="<%=ZC_STATIC_MODE%>" />
<table width='100%' style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0'>
<tr><td><p  align='left'><b>·分类页的静态化选项</b><br/><span class='note'>主机安装有ISAPI Rewrite或是URL Rewrite可以开启Rewrite选项</span></p></td><td><p><label><input type="radio" value="ACTIVE" name="STATIC" <%=IIF(ZC_STATIC_MODE="ACTIVE","checked='checked'","")%>/>&nbsp;&nbsp;1.动态页面(系统默认)</label>&nbsp;&nbsp;&nbsp;&nbsp;<label><input type="radio" value="REWRITE" name="STATIC" <%=IIF(ZC_STATIC_MODE="REWRITE","checked='checked'","")%>/>&nbsp;&nbsp;2.动态页面+Rewrite支持</label></p></td></tr>
<tr><td width='30%'><p align='left'><b>·首页的URL配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_DEFAULT_REGEX' name='edtZC_DEFAULT_REGEX' style='width:500px;' type='text' value='<%=ZC_DEFAULT_REGEX%>' /></p></td></tr>
<tr><td width='30%'><p>推荐配置</p></td><td>
<p><label onClick="changeval(6,1)"><input type="radio" name="radio6" />&nbsp;&nbsp;配置1:首页分页(默认) http://www.yourblog/default_2.html</label></p>
</td></tr>
<tr><td width='30%'><p align='left'><b>·分类页的URL配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_CATEGORY_REGEX' name='edtZC_CATEGORY_REGEX' style='width:500px;' type='text' value='<%=ZC_CATEGORY_REGEX%>' /></label></p></td></tr>
<tr><td width='30%'><p>推荐配置</p></td><td>
<p><label onClick="changeval(3,1)"><input type="radio" name="radio3" />&nbsp;&nbsp;配置1:分类ID型(默认) http://www.yourblog/category-id.html</label></p>
<p><label onClick="changeval(3,2)"><input type="radio" name="radio3" />&nbsp;&nbsp;配置2:分类ID目录型 http://www.yourblog/category/id/</label></p>
<p><label onClick="changeval(3,3)"><input type="radio" name="radio3" />&nbsp;&nbsp;配置3:分类别名目录 http://www.yourblog/categroy/alias/</label></p>
</td></tr>
<tr><td width='30%'><p align='left'><b>·作者页的URL配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_USER_REGEX' name='edtZC_USER_REGEX' style='width:500px;' type='text' value='<%=ZC_USER_REGEX%>' /></label></p></td></tr>
<tr><td width='30%'><p>推荐配置</p></td><td>
<p><label onClick="changeval(7,1)"><input type="radio" name="radio7" />&nbsp;&nbsp;配置1:作者ID型(默认) http://www.yourblog/author-1.html</label></p>
</td></tr>
<tr><td width='30%'><p align='left'><b>·TAGS页的URL配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_TAGS_REGEX' name='edtZC_TAGS_REGEX' style='width:500px;' type='text' value='<%=ZC_TAGS_REGEX%>' /></label></p></td></tr>
<tr><td width='30%'><p>推荐配置</p></td><td>
<p><label onClick="changeval(4,1)"><input type="radio" name="radio4" />&nbsp;&nbsp;配置1:Tags ID型(默认) http://www.yourblog/tags-id.html</label></p>
<p><label onClick="changeval(4,2)"><input type="radio" name="radio4" />&nbsp;&nbsp;配置2:Tags 名称型 http://www.yourblog/tags-name.html</label></p>
</td></tr>
<tr><td width='30%'><p align='left'><b>·日期页的URL配置</b><br/><span class='note'></span></p></td><td><p><input id='edtZC_DATE_REGEX' name='edtZC_DATE_REGEX' style='width:500px;' type='text' value='<%=ZC_DATE_REGEX%>' /></label></p></td></tr>
<tr><td width='30%'><p>推荐配置</p></td><td>
<p><label onClick="changeval(5,1)"><input type="radio" name="radio5" />&nbsp;&nbsp;配置1:日期型(默认) http://www.yourblog/date-2012-12.html</label></p>
<p><label onClick="changeval(5,2)"><input type="radio" name="radio5" />&nbsp;&nbsp;配置2:日期型2 http://www.yourblog/archives/2012-12.html</label></p>
<p><label onClick="changeval(5,3)"><input type="radio" name="radio5" />&nbsp;&nbsp;配置3:日期目录型 http://www.yourblog/archives/2012-12/</label></p>
</td></tr>
</table>
<p><span class="note">分类,作者,TAGS可用{%id%},{%name%}或{%alias%},分类的{%alias%}为空用name代替,作者的{%alias%}为空用name代替,TAGS的{%alias%}为URLEncode编码处理过的name,日期页可用{%date%}.</span></p>

</div>


				</div> <!-- End .content-box-content -->
				
			</div> <!-- End .content-box -->


<p><span class="star">注意:开启文章及页面和分类页的Rewrite支持选项后,请查看"ReWrite规则"并应用在主机上方能生效.</span></p>

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
		}
		if(a==2){
			a="#edtZC_PAGE_REGEX";
			if(b==1){b="{%host%}/{%alias%}.html"};
			if(b==2){b="{%host%}/{%alias%}/default.html"};
		}
		if(a==3){
			a="#edtZC_CATEGORY_REGEX";
			if(b==1){b="{%host%}/category-{%id%}.html"};
			if(b==2){b="{%host%}/category/{%id%}/default.html"};
			if(b==3){b="{%host%}/category-{%alias%}/default.html"};
		}
		if(a==4){
			a="#edtZC_TAGS_REGEX";
			if(b==1){b="{%host%}/tags-{%id%}.html"};
			if(b==2){b="{%host%}/tags-{%name%}.html"};
		}
		if(a==5){
			a="#edtZC_DATE_REGEX";
			if(b==1){b="{%host%}/date-{%date%}.html"};
			if(b==2){b="{%host%}/archives/{%date%}.html"};
			if(b==3){b="{%host%}/archives/{%date%}/default.html"};
		}
		if(a==6){
			a="#edtZC_DEFAULT_REGEX";
			if(b==1){b="{%host%}/default.html"};
		}
		if(a==7){
			a="#edtZC_USER_REGEX";
			if(b==1){b="{%host%}/author-{%id%}.html"};
		}

		$(a).val(b);
	}

	$(":radio[name='POST_STATIC']").live("click",function(){
		$("#edtZC_POST_STATIC_MODE").val($(this).val());
		if($(this).val()=="STATIC"){
			$("#edtZC_ARTICLE_REGEXT").val("{%host%}/{%post%}/{%alias%}.html");
			$("#edtZC_PAGE_REGEX").val("{%host%}/{%alias%}.html");
			$("input[name='radio'],input[name='radio2']").removeAttr("disabled");
		};
		if($(this).val()=="ACTIVE"){
			$("#edtZC_ARTICLE_REGEXT").val("{%host%}/view.asp?id={%id%}");
			$("#edtZC_PAGE_REGEX").val("{%host%}/view.asp?id={%id%}");
			$("input[name='radio'],input[name='radio2']").attr("disabled","disabled");
		};
		if($(this).val()=="REWRITE"){
			$("#edtZC_ARTICLE_REGEXT").val("{%host%}/{%post%}/{%alias%}.html");
			$("#edtZC_PAGE_REGEX").val("{%host%}/{%alias%}.html");	
			$("input[name='radio'],input[name='radio2']").removeAttr("disabled");
		};

	});


	$(":radio[name='STATIC']").live("click",function(){
		$("#edtZC_STATIC_MODE").val($(this).val());

		if($(this).val()=="ACTIVE"){
			$("#edtZC_DEFAULT_REGEX").val("{%host%}/catalog.asp");
			$("#edtZC_CATEGORY_REGEX").val("{%host%}/catalog.asp?cate={%id%}");
			$("#edtZC_USER_REGEX").val("{%host%}/catalog.asp?user={%id%}");
			$("#edtZC_TAGS_REGEX").val("{%host%}/catalog.asp?tags={%alias%}");
			$("#edtZC_DATE_REGEX").val("{%host%}/catalog.asp?date={%date%}");
			$("#edtZC_STATIC_MODE").val("ACTIVE");
			$("input[name='radio3'],input[name='radio4'],input[name='radio5'],input[name='radio6'],input[name='radio7']").attr("disabled","disabled");

		};
		if($(this).val()=="REWRITE"){
			$("#edtZC_DEFAULT_REGEX").val("{%host%}/default.html");
			$("#edtZC_CATEGORY_REGEX").val("{%host%}/category-{%id%}.html");
			$("#edtZC_USER_REGEX").val("{%host%}/author-{%id%}.html");
			$("#edtZC_TAGS_REGEX").val("{%host%}/tags-{%id%}.html");
			$("#edtZC_DATE_REGEX").val("{%host%}/date-{%date%}.html");
			$("#edtZC_STATIC_MODE").val("REWRITE");
			$("input[name='radio3'],input[name='radio4'],input[name='radio5'],input[name='radio6'],input[name='radio7']").removeAttr("disabled");
		};

	});

</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
