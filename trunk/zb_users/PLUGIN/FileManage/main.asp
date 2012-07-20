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
<!-- #include file="function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("FileManage")=False Then Call ShowError(48)
BlogTitle="文件管理"

Dim strAct,strPath,strOPath
strAct=Request.QueryString("act")
If strAct="" Then strAct="SiteFileMng"
strPath=Request.QueryString("path")
strOPath=Request.QueryString("opath")
If strPath="" Then strPath=BlogPath: strOPath=BlogPath
For Each Action_Plugin_FileManage_Initialize in Action_Plugin_FileManage_Initialize
		If Not IsEmpty(sAction_Plugin_FileManage_Initialize) Then Call Execute(sAction_Plugin_FileManage_Initialize)
Next
Select Case Request.QueryString("act")
		Case "SiteFileDownload" Call FileManage_DownloadFile(strPath)
End Select

Call SetBlogHint_Custom("‼ 提示:错误的编辑或删除系统文件会导致Blog无法运行;请保护好管理员账号,防止他人通过此功能威胁空间安全.")
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
#fileUpload{display:none;border:gray 1px solid}

</style>
<%If FileManage_CodeMirror=True Then%><link rel="stylesheet" href="../../../ZB_SYSTEM/admin/ueditor/third-party/codemirror2.15/codemirror.css"/><%End If%>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain">
  <div class="divHeader"><%=BlogTitle%></div>
  <div class="SubMenu"> 
<%= Response_Plugin_SiteFileMng_SubMenu%>
    <!--<span class="m-left m-now"><a href="main.asp">[插件后台管理页]</a> </span>--> 
  </div>
  <div id="divMain2">
    <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div><script type="text/javascript">ActiveLeftMenu("aSiteFileMng");</script>

    <%

'	If strOPath="" Then strOPath=BlogPath
'	If Not CheckRights(strAct) Then Call ShowError(6)
	Select Case Request.QueryString("act")

		Case "SiteFileMng","" Call FileManage_ExportSiteFileList(strPath,strOPath)
		Case "SiteFileEdt" Call FileManage_ExportSiteFileEdit(strPath,strOPath)
		Case "SiteFileDel" Call FileManage_DeleteSiteFile(strPath)
		Case "SiteFileUploadShow" Call FileManage_ExportSiteUpload(strPath)
		Case "SiteFileUpload" Call FileManage_Upload
		Case "SiteFileRename" Call FileManage_RenameFile(strPath,Request.QueryString("newfilename"))
		Case "SiteFilePst" Call FileManage_PostSiteFile(Request.Form("path"),Request.Form("opath"))
		Case "SiteCreateFolder" Call FileManage_CreateFolder(strPath)
		Case "Help" Call FileManage_Help
		Case "ThemeEditor" Response.Redirect "?act=SiteFileMng&path="&server.URLEncode(blogpath&"\zb_users\theme\"&zc_blog_theme)
		Case Else Response.Write "未知的命令"
	End Select
	%>
  </div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%
For Each Action_Plugin_FileManage_Terminate in Action_Plugin_FileManage_Terminate
		If Not IsEmpty(sAction_Plugin_FileManage_Terminate) Then Call Execute(sAction_Plugin_FileManage_Terminate)
Next
Call System_Terminate()

%>
