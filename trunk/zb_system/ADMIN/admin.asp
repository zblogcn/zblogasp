<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    admin.asp
'// 开始时间:    2004.07.30
'// 最后修改:    
'// 备    注:    管理页
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_event.asp" -->
<!-- #include file="../function/c_system_manage.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->
<%

Call System_Initialize()

'plugin node
For Each sAction_Plugin_Admin_Begin in Action_Plugin_Admin_Begin
	If Not IsEmpty(sAction_Plugin_Admin_Begin) Then Call Execute(sAction_Plugin_Admin_Begin)
Next

Call CheckReference("")

Dim strAct
strAct=Request.QueryString("act")

'检查权限
If Not CheckRights(strAct) Then Call ShowError(6)

Call GetCategory()
Call GetUser()

BlogTitle=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG046


'检查模板的更新,如有更新要重新加载
Dim strTemplateModified
Application.Lock
strTemplateModified=Application(ZC_BLOG_CLSID & "TEMPLATEMODIFIED")
Application.UnLock
If IsEmpty(strTemplateModified)=False Then
	If LCase(CStr(strTemplateModified))<>LCase(CStr(CheckTemplateModified)) Then
		Call ClearGlobeCache()
		Call LoadGlobeCache()
	End If
End If


%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<!--#include file="admin_header.asp"-->
	<link rel="stylesheet" href="../CSS/jquery.bettertip.css" type="text/css" media="screen">
	<script language="JavaScript" src="../script/jquery.bettertip.pack.js" type="text/javascript"></script>
	<script language="JavaScript" src="../script/jquery.textarearesizer.compressed.js" type="text/javascript"></script>
	<title><%=BlogTitle%></title>
</head>
<body>
<script type="text/javascript">
	$(function(){
		BT_setOptions({openWait:250, closeWait:0, cacheEnabled:true});
	})
</script>
			<!--#include file="admin_top.asp"-->
<div id="main">
<div class="main_right">
  <div class="yui">
    <div class="content">
    <div class="wrapper">
      <%
	Select Case Request.QueryString("act")
		Case "ArticleMng"
			If Request.QueryString("type")="Page" Then
			Call ExportPageList(Request.QueryString("page"),Request("cate"),Request("level"),Escape(Request("title")))
			Else
			Call ExportArticleList(Request.QueryString("page"),Request("cate"),Request("level"),Escape(Request("title")))
			End If
		Case "CategoryMng" Call ExportCategoryList(Request.QueryString("page"))
		Case "CommentMng" Call ExportCommentList(Request.QueryString("page"),Request("intContent"))
		Case "TrackBackMng" Call ExportTrackBackList(Request.QueryString("page"))
		Case "UserMng" Call ExportUserList(Request.QueryString("page"))
		Case "FileMng" Call ExportFileList(Request.QueryString("page"))
		Case "TagMng" Call ExportTagList(Request.QueryString("page"))
		Case "PlugInMng" Call ExportPluginMng()
		Case "SiteInfo" Call ExportSiteInfo()
		Case "AskFileReBuild" Call ExportFileReBuildAsk()
		Case "ThemeMng" Call ExportThemeMng()
	End Select
%>
    </div>
    </div>
  </div>
</div>
<!--#include file="admin_left.asp"-->
			</div>
<script>

$(document).ready(function(){ 

	//斑马线
	var tables=document.getElementsByTagName("table");
	for (var j = 0; j < tables.length; j++){

		var cells = tables[j].getElementsByTagName("tr");
		var b=false;
		cells[0].className="color1";
		for (var i = 1; i < cells.length; i++){
			if(b){
				cells[i].className="color2";
				b=false;
				cells[i].onmouseover=function(){
					this.className="color4";
				}
				cells[i].onmouseout=function(){
					this.className="color2";
				}
			}
			else{
				cells[i].className="color3";
				b=true;
				cells[i].onmouseover=function(){
					this.className="color4";
				}
				cells[i].onmouseout=function(){
					this.className="color3";
				}
			};

		};
	}

	$('textarea.resizable:not(.processed)').TextAreaResizer();
	$('iframe.resizable:not(.processed)').TextAreaResizer();

});

</script>
</body>
</html>
<%
'plugin node
For Each sAction_Plugin_Admin_End in Action_Plugin_Admin_End
	If Not IsEmpty(sAction_Plugin_Admin_End) Then Call Execute(sAction_Plugin_Admin_End)
Next

Call System_Terminate()

If Err.Number<>0 then
	Call ShowError(0)
End If
%>