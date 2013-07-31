<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../../zb_system/admin/ueditor/asp/aspincludefile.asp"-->

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
Dim FileManage_ShowAppsName__,FileManage_OpenCodeMirror,FileManage_DefaultPath___,FileManage_Return2List___
If strPath="" Then strPath=BlogPath: strOpenFolderPath=BlogPath

Dim objConfig
Set objConfig=New TConfig
objConfig.Load "FileManage"

If objConfig.Exists("FirstRun")=False Then
	objConfig.Write "ShowAppsName__","True"
	objConfig.Write "OpenCodeMirror","False"
	objConfig.Write "DefaultPath___",""
	objConfig.Write "Return2List___","True"
	objConfig.Write "FirstRun","guess"
	objConfig.Save
ElseIf objConfig.Read("FirstRun")="ok" Then 'v1.0
	objConfig.Write "Return2List___","True"
	objConfig.Write "FirstRun","guess"
	objConfig.Save
End If

FileManage_ShowAppsName__=CBool(objConfig.Read("ShowAppsName__"))
FileManage_OpenCodeMirror=CBool(objConfig.Read("OpenCodeMirror"))
FileManage_Return2List___=CBool(objConfig.Read("Return2List___"))
FileManage_DefaultPath___=CStr(objConfig.Read("DefaultPath___"))
If FileManage_ShowAppsName__=True Then
	Call Add_Action_Plugin("Action_Plugin_FileManage_ExportInformation_NotFound","FileManage_GetPluginName(""{path}"",""{f}"")")
	Call Add_Action_Plugin("Action_Plugin_FileManage_ExportInformation_NotFound","FileManage_GetThemeName(""{path}"",""{f}"")")
End If




For Each Action_Plugin_FileManage_Initialize in Action_Plugin_FileManage_Initialize
		If Not IsEmpty(sAction_Plugin_FileManage_Initialize) Then Call Execute(sAction_Plugin_FileManage_Initialize)
Next

Select Case strAct
		Case "SiteFileDownload" Call FileManage_DownloadFile(strPath)
		Case "SiteFilePst" Call FileManage_PostSiteFile(Request.Form("path"),Request.QueryString("OpenFolderPath"))
		Case "SiteFileDel" Call FileManage_DeleteSiteFile(strPath,IIf(Request.QueryString("folder")="true",True,False)):strAct="SiteFileMng"
		Case "SiteFileRename" Call FileManage_RenameFile(strPath,Request.QueryString("newfilename"),IIf(Request.QueryString("folder")="true",True,False)):strAct="SiteFileMng"
		Case "SiteFileUpload" Call FileManage_Upload
		Case "SiteCreateFolder" Call FileManage_CreateFolder(strPath,strOpenFolderPath):strAct="SiteFileMng"

End Select


Dim bolAjax
bolAjax=IIf(Request.QueryString("ajax")="True",True,False)

'Call SetBlogHint_Custom(" 若需要修改的数据>200K，请使用文件上传或FTP。")

%>
<%
If Not bolAjax Then
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<script type="text/javascript" src="jquery.history.js"></script>
<%If FileManage_OpenCodeMirror=True Then%>
<link rel="stylesheet" href="../../../ZB_SYSTEM/admin/ueditor/third-party/codemirror/codemirror.css"/>
<script language="JavaScript" type="text/javascript" src="../../../ZB_SYSTEM/admin/ueditor/third-party/codemirror/codemirror.js"></script>
<%End If%>
<!--<script language="JavaScript" type="text/javascript" src="jquery.dataTables.min.js"></script>
<link rel="stylesheet" href="jquery.dataTables.css"/>-->
<style type="text/css">
#fileUpload{display:none;}
#fileUpload #edit{background:none}
#fileUpload #edit .button{float:right}
#siteList {margin-top: -13px;}
/*table.dataTable tr.odd{background: #F4F4F4 !important;}
table.dataTable tr.even{background: #FFFFFF !important;}
tr.even td.sorting_1{background: #FFFFFF !important;}
table.dataTable tr.odd td.sorting_1{background: #F4F4F4 !important;}
table.dataTable tr.color4 {background: #ffffdd !important;}
table.dataTable tr.color4 td.sorting_1{background: #ffffdd !important;}*/

</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->

<div id="divMain">
<%
End If
%>
<div id="loading" style="display:none"><img src='../../../zb_system/image/admin/loading.gif' />Waiting</div>

<div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader"><%=BlogTitle%>&nbsp;&nbsp;<%
  If Request.QueryString("act")<>"Setting" Then%>
	  <a h='_' href="main.asp?act=Setting" title="设置"><img src="../../../zb_system/IMAGE/ADMIN/setting_tools.png" width="16" alt=""/></a>
  <%
  Else
  %>
  	  <a h='_' href="<%=Request.ServerVariables("HTTP_REFERER")%>" title="返回"><img src="images\upload.png" width="16" alt="" /></a>
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
	Select Case strAct

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
<%
If Not bolAjax Then
%>
</div>

<script type="text/javascript">
function something_others(domready){
	$(".rename_folder").mousedown(function(){
		var str=prompt("请输入新文件夹名");if(str!=null){this.href+="&folder=true&newfilename="+encodeURIComponent(str);this.click()}else{return false}
	});
	$(".rename_file").mousedown(function(){
		var str=prompt("请输入新文件名");if(str!=null){this.href+="&newfilename="+encodeURIComponent(str);this.click()}else{return false}
	});

	$("a[h='_']").click(function(){
		var This=$(this);
		var _href=This.attr("href");
		var cls=This.hasClass("delete_file")||This.hasClass("delete_folder");
		
		if(cls){
			if(!window.confirm("<%=ZC_MSG058%>")){return false;}
		}
		
		$("#loading").show();
		
		$.get(_href+(/\?/.test(_href)?"&":"?")+"ajax=True",{},function(data){
			$("#divMain").html(data);
			something_others(true);
		});
		History.pushState({rand:Math.random()},"<%=ZC_BLOG_TITLE & ZC_MSG044 & BlogTitle%>", _href); 
		
		$("#loading").hide();
		
		return false;
	});
	
	if(domready){
		bmx2table();
		if(!(($.browser.msie)&&($.browser.version)=='6.0')){
			$('input.checkbox').css("display","none");
			$('input.checkbox[value="True"]').after('<span class="imgcheck imgcheck-on"></span>');
			$('input.checkbox[value!="True"]').after('<span class="imgcheck"></span>');
		}else{
			$('input.checkbox').attr('readonly','readonly');
			$('input.checkbox').css('cursor','pointer');
			$('input.checkbox').click(function(){  if($(this).val()=='True'){$(this).val('False')}else{$(this).val('True')} })
		}

		$('span.imgcheck').click(function(){changeCheckValue(this)})

		$("#batch a").bind("click", function(){ BatchContinue();$("#batch p").html(" 操作正在进行中,请稍候......");});
	
		$(".SubMenu span.m-right").parent().css({"float":"right"});
		$("img[width='16']").each(function(){if($(this).parent().is("a")){$(this).parent().addClass("button")}});

	
		}
};
something_others(false);
/*
$(document).ready(function(){
	//$("tr").unbind("mouseover");
	$("#siteList").dataTable({ "oLanguage": {
            "sLengthMenu": "每页显示 _MENU_ 个文件",
            "sZeroRecords": "未找到",
            "sInfo": "显示 _START_ 到 _END_ 共 _TOTAL_ 个文件",
            "sInfoEmpty": "显示 0 to 到 of 共 0 个文件",
            "sInfoFiltered": "(从_MAX_条记录里过滤)",
			"sEmptyTable":"没有任何文件",
			"sSearch":"搜索",
			"oPaginate":{
				"sFirst":"首页","sLast":"尾页","sNext":"下一页","sPrevious":"上一页"
			}
        }
    });
})
*/
</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%End If%>
<%
For Each Action_Plugin_FileManage_Terminate in Action_Plugin_FileManage_Terminate
		If Not IsEmpty(sAction_Plugin_FileManage_Terminate) Then Call Execute(sAction_Plugin_FileManage_Terminate)
Next
Call System_Terminate()

%>
