<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ZBDK")=False Then Call ShowError(48)
BlogTitle=title
If Request.QueryString("act")="interface" Then
	Dim s,n,j,i
	s=LCase(Request.Form("interface"))
	
	Select Case Left(s,6)
		Case "action"
		'我X不能用Join
			Response.Write "<table width='100%'><tr><td height='40'>代码（共"
			Execute "Response.Write Ubound("&s&")"
			Response.Write "行）</td></tr>"
			Execute "j="&s
			For i=1 To Ubound(j)
				n=n & "<tr><td height='40'>" & TransferHTML(j(i),"[html-format]") & "</td></tr>"
			Next
			Response.Write n
			Response.Write "</table>"
		Case "filter"
			Response.Write "<table width='100%'><tr><td height='40'>函数（共"
			Execute "j=Split(s"&s&",""|"")"
			Response.Write Ubound(j)&"个）</td></tr>"
			For i=0 To Ubound(j)-1
				n=n & "<tr><td height='40'>" & TransferHTML(j(i),"[html-format]") & "</td></tr>"
			Next
			Response.Write n
			Response.Write "</table>"
		Case "respon"
			Execute "Response.Write TransferHTML("&s&",""[html-format]"")"
		Case Else
			Response.Write "这似乎不是Z-Blog的接口吧？"
	End Select
	
	Response.End
End If
%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<script type="text/javascript">
	var lists={
		"action":["Action_Plugin_Admin_Begin",
"Action_Plugin_Admin_End",
"Action_Plugin_ArticleDel_Begin",
"Action_Plugin_ArticleDel_Succeed",
"Action_Plugin_ArticleEdt_Begin",
"Action_Plugin_ArticleMng_Begin",
"Action_Plugin_ArticlePst_Begin",
"Action_Plugin_ArticlePst_Succeed",
"Action_Plugin_AskFileReBuild_Begin",
"Action_Plugin_Batch_Begin",
"Action_Plugin_Batch_End",
"Action_Plugin_BlogAdmin_Begin",
"Action_Plugin_BlogLogin_Begin",
"Action_Plugin_BlogLogout_Begin",
"Action_Plugin_BlogLogout_Succeed",
"Action_Plugin_BlogReBuild_Archives_Begin",
"Action_Plugin_BlogReBuild_Authors_Begin",
"Action_Plugin_BlogReBuild_Begin",
"Action_Plugin_BlogReBuild_Calendar_Begin",
"Action_Plugin_BlogReBuild_Catalogs_Begin",
"Action_Plugin_BlogReBuild_Categorys_Begin",
"Action_Plugin_BlogReBuild_Comments_Begin",
"Action_Plugin_BlogReBuild_Default_Begin",
"Action_Plugin_BlogReBuild_End",
"Action_Plugin_BlogReBuild_Functions_Begin",
"Action_Plugin_BlogReBuild_GuestComments_Begin",
"Action_Plugin_BlogReBuild_Previous_Begin",
"Action_Plugin_BlogReBuild_Statistics_Begin",
"Action_Plugin_BlogReBuild_Tags_Begin",
"Action_Plugin_BlogReBuild_TrackBacks_Begin",
"Action_Plugin_BlogVerify_Begin",
"Action_Plugin_BlogVerify_Succeed",
"Action_Plugin_BuildAllCache_Begin",
"Action_Plugin_Catalog_Begin",
"Action_Plugin_Catalog_End",
"Action_Plugin_CategoryDel_Begin",
"Action_Plugin_CategoryDel_Succeed",
"Action_Plugin_CategoryEdt_Begin",
"Action_Plugin_CategoryMng_Begin",
"Action_Plugin_CategoryPst_Begin",
"Action_Plugin_CategoryPst_Succeed",
"Action_Plugin_CheckRights_Begin",
"Action_Plugin_Command_Begin",
"Action_Plugin_Command_End",
"Action_Plugin_CommentAudit_Begin",
"Action_Plugin_CommentAudit_Success",
"Action_Plugin_CommentDelBatch_Begin",
"Action_Plugin_CommentDelBatch_Succeed",
"Action_Plugin_CommentDel_Begin",
"Action_Plugin_CommentDel_Succeed",
"Action_Plugin_CommentEdt_Begin",
"Action_Plugin_CommentMng_Begin",
"Action_Plugin_CommentPost_Begin",
"Action_Plugin_CommentPost_Succeed",
"Action_Plugin_CommentRev_Begin",
"Action_Plugin_CommentRev_Succeed",
"Action_Plugin_CommentSav_Begin",
"Action_Plugin_CommentSav_Succeed",
"Action_Plugin_Default_Begin",
"Action_Plugin_Default_End",
"Action_Plugin_DirectoryReBuild_Begin",
"Action_Plugin_DirectoryReBuild_Succeed",
"Action_Plugin_EditCatalog_Form",
"Action_Plugin_EditComment_Form",
"Action_Plugin_EditTag_Form",
"Action_Plugin_EditUser_Form",
"Action_Plugin_Edit_Begin",
"Action_Plugin_Edit_Catalog_Begin",
"Action_Plugin_Edit_Comment_Begin",
"Action_Plugin_Edit_Form",
"Action_Plugin_Edit_Link_Begin",
"Action_Plugin_Edit_Setting_Begin",
"Action_Plugin_Edit_Tag_Begin",
"Action_Plugin_Edit_UEditor_Begin",
"Action_Plugin_Edit_User_Begin",
"Action_Plugin_Edit_ueditor_getArticleInfo",
"Action_Plugin_ExportATOM_Begin",
"Action_Plugin_ExportRSS_Begin",
"Action_Plugin_Feed_Begin",
"Action_Plugin_Feed_End",
"Action_Plugin_FileDelBatch_Begin",
"Action_Plugin_FileDelBatch_Succeed",
"Action_Plugin_FileDel_Begin",
"Action_Plugin_FileDel_Succeed",
"Action_Plugin_FileMng_Begin",
"Action_Plugin_FileReBuild_Begin",
"Action_Plugin_FileReBuild_End",
"Action_Plugin_FileSnd_Begin",
"Action_Plugin_FileUpload_Begin",
"Action_Plugin_FileUpload_Succeed",
"Action_Plugin_FunctionDel_Begin",
"Action_Plugin_FunctionDel_Succeed",
"Action_Plugin_FunctionEdt_Begin",
"Action_Plugin_FunctionMng_Begin",
"Action_Plugin_FunctionSav_Begin",
"Action_Plugin_FunctionSav_Succeed",
"Action_Plugin_GetRights_Begin",
"Action_Plugin_Guestbook_Begin",
"Action_Plugin_Guestbook_End",
"Action_Plugin_LinkMng_Begin",
"Action_Plugin_LinkSav_Begin",
"Action_Plugin_LinkSav_Succeed",
"Action_Plugin_MakeBlogReBuild_Begin",
"Action_Plugin_MakeBlogReBuild_Core_Begin",
"Action_Plugin_MakeBlogReBuild_Core_End",
"Action_Plugin_MakeBlogReBuild_End",
"Action_Plugin_MakeCalendar_Begin",
"Action_Plugin_MakeFileReBuild_Begin",
"Action_Plugin_MakeFileReBuild_End",
"Action_Plugin_OpenConnect",
"Action_Plugin_PlugInActive_Begin",
"Action_Plugin_PlugInActive_Succeed",
"Action_Plugin_PlugInDisable_Begin",
"Action_Plugin_PlugInDisable_Succeed",
"Action_Plugin_PlugInMng_Begin",
"Action_Plugin_Search_Begin",
"Action_Plugin_Searching_Begin",
"Action_Plugin_Searching_End",
"Action_Plugin_SettingMng_Begin",
"Action_Plugin_SettingSav_Begin",
"Action_Plugin_SettingSav_Succeed",
"Action_Plugin_SiteFileDel_Begin",
"Action_Plugin_SiteFileDel_Succeed",
"Action_Plugin_SiteFileEdt_Begin",
"Action_Plugin_SiteFileMng_Begin",
"Action_Plugin_SiteFilePst_Begin",
"Action_Plugin_SiteFilePst_Succeed",
"Action_Plugin_SiteInfo_Begin",
"Action_Plugin_System_Initialize",
"Action_Plugin_System_Initialize_Succeed",
"Action_Plugin_System_Terminate",
"Action_Plugin_System_Terminate_WithOutDB",
"Action_Plugin_TArticleList_ExportBar_Begin",
"Action_Plugin_TArticleList_ExportBar_End'",
"Action_Plugin_TArticleList_Export_Begin",
"Action_Plugin_TArticleList_Export_End",
"Action_Plugin_TArticleList_Search_Begin",
"Action_Plugin_TArticleList_Search_End",
"Action_Plugin_TArticle_Export_Begin",
"Action_Plugin_TArticle_Export_CMTandTB_Begin",
"Action_Plugin_TArticle_Export_CommentPost_Begin",
"Action_Plugin_TArticle_Export_End",
"Action_Plugin_TArticle_Export_Mutuality_Begin",
"Action_Plugin_TArticle_Export_NavBar_Begin",
"Action_Plugin_TArticle_Export_Tag_Begin",
"Action_Plugin_TArticle_Save_Begin",
"Action_Plugin_TArticle_Url",
"Action_Plugin_TCategory_Url",
"Action_Plugin_TComment_Avatar",
"Action_Plugin_TTag_Url",
"Action_Plugin_TUser_Url",
"Action_Plugin_TagDel_Begin",
"Action_Plugin_TagDel_Succeed",
"Action_Plugin_TagEdt_Begin",
"Action_Plugin_TagMng_Begin",
"Action_Plugin_TagPst_Begin",
"Action_Plugin_TagPst_Succeed",
"Action_Plugin_Tags_Begin",
"Action_Plugin_Tags_End",
"Action_Plugin_ThemeMng_Begin",
"Action_Plugin_ThemeSav_Begin",
"Action_Plugin_Themesav_Succeed",
"Action_Plugin_TrackBackDelBatch_Begin",
"Action_Plugin_TrackBackDelBatch_Succeed",
"Action_Plugin_TrackBackDel_Begin",
"Action_Plugin_TrackBackDel_Succeed",
"Action_Plugin_TrackBackMng_Begin",
"Action_Plugin_TrackBackPost_Begin",
"Action_Plugin_TrackBackPost_Succeed",
"Action_Plugin_TrackBackSnd_Begin",
"Action_Plugin_TrackBackSnd_Succeed",
"Action_Plugin_TrackBackUrlGet_Begin",
"Action_Plugin_UEditor_Config_Begin",
"Action_Plugin_UEditor_Config_End",
"Action_Plugin_UEditor_FileUpload_Begin",
"Action_Plugin_UEditor_FileUpload_End",
"Action_Plugin_UEditor_getRemoteImage_Begin",
"Action_Plugin_UEditor_getRemoteImage_End",
"Action_Plugin_UEditor_getcontent_Begin",
"Action_Plugin_UEditor_getcontent_End",
"Action_Plugin_UEditor_getmovie_Begin",
"Action_Plugin_UEditor_getmovie_End",
"Action_Plugin_UEditor_imageManager_Begin",
"Action_Plugin_UEditor_imageManager_End",
"Action_Plugin_UserCrt_Begin",
"Action_Plugin_UserCrt_Succeed",
"Action_Plugin_UserDel_Begin",
"Action_Plugin_UserDel_Succeed",
"Action_Plugin_UserEdt_Begin",
"Action_Plugin_UserMng_Begin",
"Action_Plugin_UserMod_Begin",
"Action_Plugin_UserMod_Succeed",
"Action_Plugin_ViewRights_Begin",
"Action_Plugin_View_Begin",
"Action_Plugin_View_End",
"Action_Plugin_Wap_Begin",
"Action_Plugin_Wap_End",
"Action_Plugin_XMLRPC_Begin",
"Action_Plugin_XMLRPC_End",
],
"filter":["Filter_Plugin_CommentAduit_Core",
"Filter_Plugin_EditUser_Core",
"Filter_Plugin_EditUser_Succeed",
"Filter_Plugin_PostArticle_Core",
"Filter_Plugin_PostArticle_Succeed",
"Filter_Plugin_PostCategory_Core",
"Filter_Plugin_PostCategory_Succeed",
"Filter_Plugin_PostComment_Core",
"Filter_Plugin_PostComment_Succeed",
"Filter_Plugin_PostTag_Core",
"Filter_Plugin_PostTag_Succeed",
"Filter_Plugin_PostTrackBack_Core",
"Filter_Plugin_PostTrackBack_Succeed",
"Filter_Plugin_TArticleList_Build_Template",
"Filter_Plugin_TArticleList_Build_TemplateTags",
"Filter_Plugin_TArticleList_Build_Template_Succeed",
"Filter_Plugin_TArticleList_Export",
"Filter_Plugin_TArticleList_ExportByCache",
"Filter_Plugin_TArticleList_ExportByMixed",
"Filter_Plugin_TArticleList_Export_TemplateTags",
"Filter_Plugin_TArticle_Build_Template",
"Filter_Plugin_TArticle_Build_TemplateTags",
"Filter_Plugin_TArticle_Build_Template_Succeed",
"Filter_Plugin_TArticle_Del",
"Filter_Plugin_TArticle_Export_Template",
"Filter_Plugin_TArticle_Export_TemplateTags",
"Filter_Plugin_TArticle_Export_Template_Sub",
"Filter_Plugin_TArticle_LoadInfoByArray",
"Filter_Plugin_TArticle_LoadInfobyID",
"Filter_Plugin_TArticle_Post",
"Filter_Plugin_TArticle_Url",
"Filter_Plugin_TCategory_Del",
"Filter_Plugin_TCategory_LoadInfoByArray",
"Filter_Plugin_TCategory_LoadInfoByID",
"Filter_Plugin_TCategory_Post",
"Filter_Plugin_TCategory_Url",
"Filter_Plugin_TComment_Del",
"Filter_Plugin_TComment_LoadInfoByArray",
"Filter_Plugin_TComment_LoadInfoByID",
"Filter_Plugin_TComment_MakeTemplate_Template",
"Filter_Plugin_TComment_MakeTemplate_TemplateTags",
"Filter_Plugin_TComment_Post",
"Filter_Plugin_TFunction_Del",
"Filter_Plugin_TFunction_LoadInfoByArray",
"Filter_Plugin_TFunction_LoadInfoByID",
"Filter_Plugin_TFunction_Post",
"Filter_Plugin_TTag_Del",
"Filter_Plugin_TTag_LoadInfoByArray",
"Filter_Plugin_TTag_LoadInfoByID",
"Filter_Plugin_TTag_MakeTemplate_Template",
"Filter_Plugin_TTag_MakeTemplate_TemplateTags",
"Filter_Plugin_TTag_Post",
"Filter_Plugin_TTag_Url",
"Filter_Plugin_TTrackBack_Del",
"Filter_Plugin_TTrackBack_LoadInfoByArray",
"Filter_Plugin_TTrackBack_LoadInfoByID",
"Filter_Plugin_TTrackBack_MakeTemplate_Template",
"Filter_Plugin_TTrackBack_MakeTemplate_TemplateTags",
"Filter_Plugin_TTrackBack_Post",
"Filter_Plugin_TUpLoadFile_Del",
"Filter_Plugin_TUpLoadFile_LoadInfoByArray",
"Filter_Plugin_TUpLoadFile_LoadInfoByID",
"Filter_Plugin_TUpLoadFile_UpLoad",
"Filter_Plugin_TUser_Del",
"Filter_Plugin_TUser_Edit",
"Filter_Plugin_TUser_LoadInfoByArray",
"Filter_Plugin_TUser_LoadInfobyID",
"Filter_Plugin_TUser_Register",
"Filter_Plugin_TUser_Url",
"Filter_Plugin_UEditor_Config",
]
,"respon":["Response_Plugin_Admin_Footer",
"Response_Plugin_Admin_Header",
"Response_Plugin_Admin_Js_Add",
"Response_Plugin_Admin_Left",
"Response_Plugin_Admin_SiteInfo",
"Response_Plugin_Admin_Top",
"Response_Plugin_ArticleEdt_SubMenu",
"Response_Plugin_ArticleMng_SubMenu",
"Response_Plugin_AskFileReBuild_SubMenu",
"Response_Plugin_CategoryEdt_SubMenu",
"Response_Plugin_CategoryMng_SubMenu",
"Response_Plugin_CommentEdt_SubMenu",
"Response_Plugin_CommentMng_SubMenu",
"Response_Plugin_EditCatalog_Form",
"Response_Plugin_EditComment_Form",
"Response_Plugin_EditTag_Form",
"Response_Plugin_EditUser_Form",
"Response_Plugin_Edit_Form",
"Response_Plugin_Edit_Form2",
"Response_Plugin_Edit_Form3",
"Response_Plugin_Edit_UbbTag",
"Response_Plugin_FileMng_SubMenu",
"Response_Plugin_FunctionMng_SubMenu",
"Response_Plugin_Function_SubMenu",
"Response_Plugin_Html_Js_Add",
"Response_Plugin_LinkMng_SubMenu",
"Response_Plugin_PlugInMng_SubMenu",
"Response_Plugin_SettingMng_SubMenu",
"Response_Plugin_SiteFileEdt_SubMenu",
"Response_Plugin_SiteFileMng_SubMenu",
"Response_Plugin_SiteInfo_SubMenu",
"Response_Plugin_TagEdt_SubMenu",
"Response_Plugin_TagMng_SubMenu",
"Response_Plugin_ThemeMng_SubMenu",
"Response_Plugin_TrackBackMng_SubMenu",
"Response_Plugin_UserEdt_SubMenu",
"Response_Plugin_UserMng_SubMenu"]
		}
		
function showlist()
{

		var str="",p=lists[$("#type").val()];
		for(var i=0; i<=p.length-1; i++){
			var o=p[i];
			str+="<option value='"+o+"'>"+o+"</option>"
		}
		$("#list").html(str);
		return str
	}
</script>
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"> <%=ZBDK.submenu(4)%> </div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("zbdk");</script>
            <form id="form1" onSubmit="return false">
            <label for="interface">输入接口名</label>
            <input type="text" name="interface" id="interface" style="width:80%"/>
            <input type="submit" name="ok" id="ok" value="查看" onClick=""/>
            <p>或选择接口名：<select name="type" id="type" onclick="showlist()">
            <option value="action">Action</option>
            <option value="filter">Filter</option>
            <option value="respon">Response</option>
            </select><select name="list" id="list" style="width:80%" onclick="$('#interface').val($(this).val())"></select></p>
            </form>
           
            <div id="result"></div>
          </div>
        </div>
        <script type="text/javascript">
		$(document).ready(function() {
			showlist();
            $("#form1").bind("submit",function(){
				$("#result").html("Waiting...");
				$.post("main.asp?act=interface",{"interface":$("#interface").val()},function(data){
					$("#result").html(data);
					bmx2table();
				}
				)
			}
			)
        });
		</script>
        <!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->