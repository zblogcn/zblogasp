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

Call GetFunction()

'检查读者墙列表=============================================
Function FunctionSav(FunctionContent)
Set objConfig=New TConfig
objConfig.Load("X2013")
if FunctionMetas.GetValue("x2013_showlink")=Empty Then
	Set objFunction=New TFunction
	objFunction.ID=0
	objFunction.Name="底部导航连接"
	objFunction.FileName="x2013_showlink"
	objFunction.HtmlID="x2013_ShowLink"
	objFunction.Ftype="ul"
	objFunction.MaxLi=6
Else
	Set objFunction=Functions(FunctionMetas.GetValue("x2013_showlink"))
End if
	objFunction.IsSystem=False
	objFunction.Content= FunctionContent
	objFunction.save
	Call SaveFunctionType()
	Call MakeBlogReBuild_Core()
End Function

Dim strAct,FctContent,objConfig,objFunction
Set objFunction=New TFunction
strAct=Request.QueryString("act")
If strAct="SaveFct" Then
	FctContent=Request.Form("inpContent")
	FunctionSav(FctContent)
	Call SetBlogHint(True,Empty,True)
Else
	Set objFunction=Functions(FunctionMetas.GetValue("x2013_showlink"))
	FctContent=Replace(objFunction.Content,"<#ZC_BLOG_HOST#>",BlogHost)
End If

%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<style>
p{line-height:1.5em;padding:0.5em 0;}
</style>
<link href="modern.css" rel="stylesheet">
<script type="text/javascript" src="accordion.js"></script>
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
<script type="text/javascript">ActiveTopMenu("aX2013");</script> 
<div id="divMain">
	<div id="ShowBlogHint"><%Call GetBlogHint()%></div>
	<!--<div class="divHeader"><%=BlogTitle%></div>-->

<div class="SubMenu"><a href="main.asp"><span class="m-left"><%=BlogTitle%></span></a><a href="navbar.asp"><span class="m-left m-now">导航管理</span></a><a href="about.asp"><span class="m-left">主题说明</span></a></div>
	<div id="divMain2">
		<div class="widget-list ui-droppable" style="min-width: 830px;">
		<div class="widget-list-header">添加导航链接</div>
		<div class="widget-list-note">请选择您要添加的链接类型</div>
			<ul data-role="accordion" class="accordion span10">
                    <li>
                        <a href="#">文章</a>
                        <div>
                           <div class="input-control select">
                                    <select multiple="1" size="10" id="post">
										<%Response.Write GetContent("Post")%>
                                    </select>
                                </div>
								<input type="button" value="添加" onclick="addsekectpartent('post')"/>
                        </div>
                    </li>
                    <li class="">
                        <a href="#">独立页面</a>
                        <div style="display: none;">
                           <div class="input-control select">
                                    <select multiple="1" size="8" id="page">
										<%Response.Write GetContent("Page")%>
                                    </select>
                                </div>
								<input type="button" value="添加" onclick="addsekectpartent('page')"/>
                        </div>
                    </li>
                    <li>
                        <a href="#">分类</a>
                        <div>
                            <div class="input-control select">
                                    <select multiple="1" size="8" id="cate">
										<%Response.Write GetContent("Category")%>
                                    </select>
                                </div>
								<input type="button" value="添加" onclick="addsekectpartent('cate')"/>
                        </div>
                    </li>
					<li>
                        <a href="#">Tags标签</a>
                        <div>
                            <div class="input-control select">
                                    <select multiple="1" size="10" id="tag">
										<%Response.Write GetContent("Tags")%>
                                    </select>
                                </div>
								<input type="button" value="添加" onclick="addsekectpartent('tag')"/>
                        </div>
                    </li>
                    <li>
                        <a href="#">自定义链接</a>
                        <div>
                            <div class="input-control text">
								<input type="url" id="addurl" placeholder="输入网址" style="width:50%"><input type="text"  id="addtitle" placeholder="输入标题" style="width:20%"><input type="button" value="添加" onclick="addurlpartent()">
							 </div>
                        </div>
                    </li>
                </ul>

		</div>

		<div class="siderbar-list">
			<div class="siderbar-drop" id="siderbar">
			<div class="siderbar-header">拖动链接进行排序</div>
			<div  class="siderbar-sort-list ui-sortable">
			<div class="widget widget_source_other ">
			<div class="page-sidebar">
			  <ul id="connect"><%Response.Write FctContent%></ul>
			</div>
			<div class="clear"></div>
				  <div id="result">
				  <form id="form1" name="form1" method="post" action="?act=SaveFct"  onsubmit="return verify()">
				  <div style="display:none;"><textarea name="inpContent" id="inpContent"></textarea></div>
				  <input type="submit" name="button" id="btn" value="确认修改"/>
				  </form>
				  </div>
			</div>
			</div></div>
		</div>
	</div>
	<div class="clear"></div>
	
<script type="text/javascript">
	$(document).ready(function(){
		$("#connect").sortable();
		$("#connect").disableSelection();

		$("#connect li").live('mouseenter mouseleave', function(event) {
		  if (event.type == 'mouseenter') {
			$(this).append("<span class='del icon-cancel' onclick='del(this)' title='删除'></span>");
		  } else {
			$(this).find(".del").remove();
		  }
		});

	});

	function del(item){
		$(item).parent().remove();
	}

	function verify(){
		var result = document.getElementById("connect").innerHTML;
		//alert(result);
		document.getElementById("inpContent").value=result;
		//document.getElementById("form1").action="?act=Save";
		return true
	}
	function addsekectpartent(vartype){
		var result = document.getElementById("connect").innerHTML;

		if(vartype=="post"){
			$("#post option:selected").each(function() {
				result = result+"<li class='menu-item'><a href='"+$(this).val()+"'>"+$(this).text()+"</a></li>";
			});
		}else if(vartype=="page"){
			$("#page option:selected").each(function() {
				result = result+"<li class='menu-item'><a href='"+$(this).val()+"'>"+$(this).text()+"</a></li>";
			});
		}else if(vartype=="cate"){
			$("#cate option:selected").each(function() {
				result = result+"<li class='menu-item'><a href='"+$(this).val()+"'>"+$(this).text()+"</a></li>";
			});
		}else if(vartype=="tag"){
			$("#tag option:selected").each(function() {
				result = result+"<li class='menu-item'><a href='"+$(this).val()+"'>"+$(this).text()+"</a></li>";
			});
		}
		//alert(result);
		document.getElementById("connect").innerHTML = result;
	}
	function addurlpartent(){
		var addurl = document.getElementById("addurl").value;
		var addtitle = document.getElementById("addtitle").value;
		var result = document.getElementById("connect").innerHTML;
		result = result+"<li class='menu-item'><a href='"+addurl+"'>"+addtitle+"</a></li>";
		document.getElementById("connect").innerHTML = result;
	}

</script>	


</div>
<!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
