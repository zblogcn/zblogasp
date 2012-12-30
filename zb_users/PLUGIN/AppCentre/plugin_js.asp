<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_admin_js.asp
'// 开始时间:    
'// 最后修改:    
'// 备    注:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Response.ContentType="application/x-javascript" %>
<!-- #include file="../../../zb_users/c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../../../zb_users/plugin/p_config.asp" -->

<% Response.Clear %>

$(document).ready(function(){ 

$("#divMain2").prepend("<form class='search' name='edit' id='edit' method='post' enctype='multipart/form-data' action='"+bloghost+"zb_users/plugin/appcentre/app_upload.asp'><p>上传插件zba文件:&nbsp;<input type='file' id='edtFileLoad' name='edtFileLoad' size='40' />&nbsp;&nbsp;&nbsp;&nbsp;<input type='submit' class='button' value='提交' name='B1' />&nbsp;&nbsp;<input class='button' type='reset' value='重置' name='B2' />&nbsp;</p></form>")


$("tr").each(function(){
	$(this).append("<td width='10%' align='center'></td>");
});

$("tr").first().children().last().append("AppCentre");

$(".plugin").each(function(){
	var t=$(this).find("strong").html();
	var s=""
	s=s+"<a href='"+bloghost+"zb_users/plugin/appcentre/plugin_edit.asp?id="+t+"' title='编辑该插件信息'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/application_edit.png'/></a>";
	s=s+"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='"+bloghost+"zb_users/plugin/appcentre/plugin_pack.asp?id="+t+"' title='导出该插件' target='_blank'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/download.png'/></a>";
	//s=s+"&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' onclick='checkApp(\""+t+"\")' title='更新该插件'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/refresh.png'/></a>";
	if(!$(this).hasClass("plugin-on")){
		$(this).parent().children().last().prev().append("&nbsp;&nbsp;&nbsp;&nbsp;<a href='"+bloghost+"zb_users/plugin/appcentre/plugin_del.asp?id="+t+"' title='删除该插件' onclick='return window.confirm(\"单击“确定”继续。单击“取消”停止。\");'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/delete.png'/></a>");
	};
	$(this).parent().children().last().append(s);
});



});

function checkApp(id){
$.get(bloghost+"zb_users/plugin/appcentre/plugin_update.asp?id="+id,function(data){alert(data);});
}