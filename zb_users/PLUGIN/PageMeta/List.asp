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
<!-- #include file="../../../ZB_SYSTEM/function/c_system_manage.asp" -->
<!-- #include file="function.asp"-->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("PageMeta")=False Then Call ShowError(48)
BlogTitle="PageMeta"

Dim Action
Select Case Request.QueryString("act")
	Case "ArticleMng"	Action=1
	Case "CategoryMng"  Action=2
	Case "UserMng" Action=3
	Case "TagMng" Action=4
End Select
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"><div id="ShowBlogHint"><%Call GetBlogHint()%></div>
      
    
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
<%=PageMeta_ExportBar(Action)%>
  </div>
  <div id="divMain2">
 <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
    <%
Select Case Action
Case 1
		If Request.QueryString("type")="Page" Then
		Call PageMeta_ExportPageList(Request.QueryString("page"),Request("cate"),Request("level"),Escape(Request("title")))
		Else
		Call PageMeta_ExportArticleList(Request.QueryString("page"),Request("cate"),Request("level"),Escape(Request("title")))
		End If
Case 2
	Call PageMeta_ExportCategoryList(Request.QueryString("page"))
Case 3
	Call PageMeta_ExportUserList(Request.QueryString("page"))
Case 4
	Call PageMeta_ExportTagList(Request.QueryString("page"))
End Select
	

	%>

  </div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
Call System_Terminate()

%>
