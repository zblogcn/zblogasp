<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("FileManage")=False Then Call ShowError(48)
BlogTitle="文件管理"
Set FileManage_FSO=Server.CreateObject("Scripting.FileSystemObject")
Dim strAct,strPath,strOpenFolderPath
strAct=Request.QueryString("act")
If strAct="" Then strAct="SiteFileMng"
strPath=Request.QueryString("path")
strOpenFolderPath=Request.QueryString("OpenFolderPath")
Dim FileManage_ShowAppsName__,FileManage_OpenCodeMirror,FileManage_DefaultPath___
If strPath="" Then strPath=BlogPath: strOpenFolderPath=BlogPath

Dim objConfig
Set objConfig=New TConfig
objConfig.Load "FileManage"

If objConfig.Exists("FirstRun")=False Then
	objConfig.Write "ShowAppsName__","True"
	objConfig.Write "OpenCodeMirror","False"
	objConfig.Write "DefaultPath___",""
	objConfig.Write "FirstRun","ok"
	objConfig.Save
End If

FileManage_ShowAppsName__=CBool(objConfig.Read("ShowAppsName__"))
FileManage_OpenCodeMirror=CBool(objConfig.Read("OpenCodeMirror"))
FileManage_DefaultPath___=CStr(objConfig.Read("DefaultPath___"))
If FileManage_ShowAppsName__=True Then
	Call Add_Action_Plugin("Action_Plugin_FileManage_ExportInformation_NotFound","FileManage_GetPluginName(""{path}"",""{f}"")")
	Call Add_Action_Plugin("Action_Plugin_FileManage_ExportInformation_NotFound","FileManage_GetThemeName(""{path}"",""{f}"")")
End If

For Each Action_Plugin_FileManage_Initialize in Action_Plugin_FileManage_Initialize
		If Not IsEmpty(sAction_Plugin_FileManage_Initialize) Then Call Execute(sAction_Plugin_FileManage_Initialize)
Next

Select Case Request.QueryString("act")
		Case "SiteFileDownload" Call FileManage_DownloadFile(strPath)
		Case "SiteFilePst" Call FileManage_PostSiteFile(Request.Form("path"),Request.QueryString("OpenFolderPath"))
		Case "SiteFileDel" Call FileManage_DeleteSiteFile(strPath,IIf(Request.QueryString("folder")="true",True,False))
		Case "SiteFileRename" Call FileManage_RenameFile(strPath,Request.QueryString("newfilename"),IIf(Request.QueryString("folder")="true",True,False))
		Case "SiteFileUpload" Call FileManage_Upload
		Case "SiteCreateFolder" Call FileManage_CreateFolder(strPath,strOpenFolderPath)

End Select

'Call SetBlogHint_Custom(" 若需要修改的数据>200K，请使用文件上传或FTP。")

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
#fileUpload{display:none;border:gray 1px solid}

</style>
<%If FileManage_OpenCodeMirror=True Then%>

<link rel="stylesheet" href="../../../ZB_SYSTEM/admin/ueditor/third-party/codemirror/codemirror.css"/>
<script language="JavaScript" type="text/javascript" src="../../../ZB_SYSTEM/admin/ueditor/third-party/codemirror/codemirror.js"></script>
<%End If%>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain"> <div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%>&nbsp;&nbsp;<%
  If Request.QueryString("act")<>"Setting" Then%>
	  <a href="main.asp?act=Setting" title="设置"><img src="../../../zb_system/IMAGE/ADMIN/setting_tools.png"/></a>
  <%
  Else
  %>
  	  <a href="javascript:history.back(-1)" title="返回"><img src="images\upload.png"/></a>
  <%
  End If
  %>
</div>
  <div class="SubMenu"> 
<%= Response_Plugin_SiteFileMng_SubMenu%>
    <!--<span class="m-left m-now"><a href="main.asp">[插件后台管理页]</a> </span>--> 
  </div>
  <div id="divMain2">
   <script type="text/javascript">ActiveLeftMenu("aSiteFileMng");</script>

    <%

'	If strOpenFolderPath="" Then strOpenFolderPath=BlogPath
'	If Not CheckRights(strAct) Then Call ShowError(6)
	Select Case Request.QueryString("act")

		Case "" Call FileManage_ExportSiteFileList(BlogPath & FileManage_DefaultPath___,"")
		Case "SiteFileMng" Call FileManage_ExportSiteFileList(strPath,strOpenFolderPath)
		Case "SiteFileEdt" Call FileManage_ExportSiteFileEdit(strPath,strOpenFolderPath,IIf(Request.QueryString("charset")=Empty,"",Request.QueryString("charset")))
		Case "SiteFileUploadShow" Call FileManage_ExportSiteUpload(strPath)
		Case "ThemeEditor" Response.Redirect "?act=SiteFileMng&path="&server.URLEncode(blogpath&"zb_users\theme\"&zc_blog_theme)
		Case "Setting" Call FileManage_Setting
		Case "SaveSetting" Call FileManage_SaveSetting
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
