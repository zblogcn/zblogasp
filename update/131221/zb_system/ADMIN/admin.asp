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
<% On Error Resume Next %>
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

'检查模板的更新,如有更新要重新加载
Dim strTemplateModified
Application.Lock
strTemplateModified=Application(ZC_BLOG_CLSID & "TEMPLATEMODIFIED")
Application.UnLock
If IsEmpty(strTemplateModified)=False Then
	If LCase(CStr(strTemplateModified))<>LCase(CStr(CheckTemplateModified)) Then
		Call BlogReBuild_Default()
	End If
End If

Dim act
act=Request.QueryString("act")
if act="" Then act="SiteInfo"

'plugin node
For Each sAction_Plugin_Admin_Begin in Action_Plugin_Admin_Begin
	If Not IsEmpty(sAction_Plugin_Admin_Begin) Then Call Execute(sAction_Plugin_Admin_Begin)
Next

'检查权限
If Not CheckRights(act) Then Call ShowError(6)

BlogTitle=ZC_MSG022

%>
<!--#include file="admin_header.asp"-->
<!--#include file="admin_top.asp"-->
    <div id="divMain">
<%	Call GetBlogHint()	%>
      <%
	Select Case act
		Case "ArticleMng"
			If Request.QueryString("type")="Page" Then
			Call ExportPageList(Request.QueryString("page"),Request("cate"),Request("level"),Escape(Request("title")))
			Else
			Call ExportArticleList(Request.QueryString("page"),Request("cate"),Request("level"),Request("istop"),Escape(Request("title")))
			End If
		Case "CategoryMng" Call ExportCategoryList(Request.QueryString("page"))
		Case "CommentMng" Call ExportCommentList(Request.QueryString("page"),Request("intContent"),Request("isCheck"))
		Case "TrackBackMng" Call ExportTrackBackList(Request.QueryString("page"))
		Case "UserMng" Call ExportUserList(Request.QueryString("page"))
		Case "FileMng" Call ExportFileList(Request.QueryString("page"))
		Case "TagMng" Call ExportTagList(Request.QueryString("page"))
		Case "PlugInMng" Call ExportPluginMng()
		Case "SiteInfo" Call ExportSiteInfo()
		Case "AskFileReBuild" Call ExportFileReBuildAsk()
		Case "ThemeMng" Call ExportThemeMng()
		Case "FunctionMng" Call ExportFunctionList()
		Case Else Call ExportSiteInfo()
	End Select
%>
    </div>
<!--#include file="admin_footer.asp"-->
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
<!-- <%=RunTime()%>ms -->