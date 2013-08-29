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
<!-- #include file="function.asp"-->
<% 

Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level=1 Then
	If CheckPluginState("AppCentre")=True Then 
Call LoadPluginXmlInfo("AppCentre")
Call AppCentre_InitConfig
	End If
End If
Response.Clear

%>

$(document).ready(function(){ 

$("#divMain2").prepend("<form class='search' name='edit' id='edit' method='post' enctype='multipart/form-data' action='"+bloghost+"zb_users/plugin/appcentre/app_upload.asp'><p>本地上传主题zba文件:&nbsp;<input type='file' id='edtFileLoad' name='edtFileLoad' size='40' />&nbsp;&nbsp;&nbsp;&nbsp;<input type='submit' class='button' value='提交' name='B1' />&nbsp;&nbsp;<input class='button' type='reset' value='重置' name='B2' />&nbsp;</p></form>");



$(".theme").each(function(){
	var t=$(this).find("strong").html();
	var s="<p>";
<%If enable_develop="True" Then%>
	s=s+"<a href='"+bloghost+"zb_users/plugin/appcentre/theme_edit.asp?id="+t+"' title='编辑该主题信息'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/application_edit.png'/></a>";

	if($(this).hasClass("theme-now")){
		s=s+"&nbsp;&nbsp;&nbsp;&nbsp;<a href='"+bloghost+"zb_system/admin/edit_function.asp?source=theme_"+t+"' title='给该主题增加侧栏模块'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/bricks.png'/></a>";
	}

	s=s+"&nbsp;&nbsp;&nbsp;&nbsp;<a href='"+bloghost+"zb_users/plugin/appcentre/theme_pack.asp?id="+t+"' title='导出该主题' target='_blank'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/download.png'/></a>";

	//if(disableupdatetheme.search(t + ":")==-1){
	//	s=s+"&nbsp;&nbsp;&nbsp;&nbsp;<a href='"+bloghost+"zb_users/plugin/appcentre/checkupdate.asp?act=dut&id="+t+"' title='禁止应用中心更新该主题'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/refresh.png'/></a>";
	//}else{
	//	s=s+"&nbsp;&nbsp;&nbsp;&nbsp;<a href='"+bloghost+"zb_users/plugin/appcentre/checkupdate.asp?act=eut&id="+t+"' title='允许应用中心更新该主题'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/refresh2.png'/></a>";
	//}

<%End If%>
	if($(this).hasClass("theme-other")){
		s=s+"&nbsp;&nbsp;&nbsp;&nbsp;"
		s=s+"<a href='"+bloghost+"zb_users/plugin/appcentre/theme_del.asp?id="+t+"' title='删除该主题' onclick='return window.confirm(\"单击“确定”继续。单击“取消”停止。\");'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/delete.png'/></a>";
	}

<%If enable_develop="True" Then%>
<%If login_pw<>"" Then%>
	s=s+"&nbsp;&nbsp;&nbsp;&nbsp;<a href='"+bloghost+"zb_users/plugin/appcentre/submit.asp?type=theme&amp;id="+t+"' title='上传主题到官方网站应用中心' target='_blank'><img height='16' width='16' src='"+bloghost+"zb_users/plugin/appcentre/images/drive-upload.png'/></a>";
<%End If%>
<%End If%>
	s=s+"</p>";
	$(this).append(s);
	
});

});

function checkApp(id){
$.get(bloghost+"zb_users/plugin/appcentre/theme_update.asp?id="+id,function(data){alert(data);});
}