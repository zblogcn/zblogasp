<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    
'// 版权所有:    
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    plugin.asp
'// 开始时间:    2007.11.28
'// 最后修改:    
'// 备    注:    插件页
'///////////////////////////////////////////////////////////////////////////////
%>
<%
'接口分三大类
'分别对应3个方法
'加入接口时请调用这几个方法


'1.action
'行为动作
'调用过程为Call Add_Action_Plugin("plugname","actioncode")

'2.filter
'过滤器
'调用过程为Call Add_Filter_Plugin("plugname","functionname")

'3.response
'纯输出
'调用过程为Call Add_Response_Plugin("plugname","inputstring")



'接口说明请勿改动,为了程序自动生成WIKI使用

'***************
'1.action
'***************



'**************************************************<
'类型:action
'名称:Action_Plugin_System_Initialize
'参数:无
'说明:在系统初始化时被调用
'**************************************************>
Dim Action_Plugin_System_Initialize()
ReDim Action_Plugin_System_Initialize(0)
Dim bAction_Plugin_System_Initialize
Dim sAction_Plugin_System_Initialize



'**************************************************<
'类型:action
'名称:Action_Plugin_System_Initialize_Succeed
'参数:无
'说明:在系统初始化成功时被调用
'**************************************************>
Dim Action_Plugin_System_Initialize_Succeed()
ReDim Action_Plugin_System_Initialize_Succeed(0)
Dim bAction_Plugin_System_Initialize_Succeed
Dim sAction_Plugin_System_Initialize_Succeed



'**************************************************<
'类型:action
'名称:Action_Plugin_System_Terminate
'参数:无
'说明:在系统终结时被调用
'**************************************************>
Dim Action_Plugin_System_Terminate()
ReDim Action_Plugin_System_Terminate(0)
Dim bAction_Plugin_System_Terminate
Dim sAction_Plugin_System_Terminate



'**************************************************<
'类型:action
'名称:Action_Plugin_System_Initialize_WithOutDB
'参数:无
'说明:在系统初始化时被调用_WithOutDB
'**************************************************>
Dim Action_Plugin_System_Initialize_WithOutDB()
ReDim Action_Plugin_System_Initialize_WithOutDB(0)
Dim bAction_Plugin_System_Initialize_WithOutDB
Dim sAction_Plugin_System_Initialize_WithOutDB




'**************************************************<
'类型:action
'名称:Action_Plugin_System_Initialize_WithOutDB_Succeed
'参数:无
'说明:在系统初始化成功时被调用_WithOutDB
'**************************************************>
Dim Action_Plugin_System_Initialize_WithOutDB_Succeed()
ReDim Action_Plugin_System_Initialize_WithOutDB_Succeed(0)
Dim bAction_Plugin_System_Initialize_WithOutDB_Succeed
Dim sAction_Plugin_System_Initialize_WithOutDB_Succeed




'**************************************************<
'类型:action
'名称:Action_Plugin_System_Terminate_WithOutDB
'参数:无
'说明:在系统终结时被调用_WithOutDB
'**************************************************>
Dim Action_Plugin_System_Terminate_WithOutDB()
ReDim Action_Plugin_System_Terminate_WithOutDB(0)
Dim bAction_Plugin_System_Terminate_WithOutDB
Dim sAction_Plugin_System_Terminate_WithOutDB




'**************************************************<
'类型:action
'名称:Action_Plugin_OpenConnect
'参数:无
'说明:OpenConnect
'**************************************************>
Dim Action_Plugin_OpenConnect()
ReDim Action_Plugin_OpenConnect(0)
Dim bAction_Plugin_OpenConnect
Dim sAction_Plugin_OpenConnect




'**************************************************<
'类型:action
'名称:Action_Plugin_Command_Begin
'参数:无
'说明:cmd.asp begin
'**************************************************>
Dim Action_Plugin_Command_Begin()
ReDim Action_Plugin_Command_Begin(0)
Dim bAction_Plugin_Command_Begin
Dim sAction_Plugin_Command_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_Command_End
'参数:无
'说明:cmd.asp end
'**************************************************>
Dim Action_Plugin_Command_End()
ReDim Action_Plugin_Command_End(0)
Dim bAction_Plugin_Command_End
Dim sAction_Plugin_Command_End



'**************************************************<
'类型:action
'名称:Action_Plugin_Admin_Begin
'参数:无
'说明:admin.asp begin
'**************************************************>
Dim Action_Plugin_Admin_Begin()
ReDim Action_Plugin_Admin_Begin(0)
Dim bAction_Plugin_Admin_Begin
Dim sAction_Plugin_Admin_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_Admin_End
'参数:无
'说明:admin.asp end
'**************************************************>
Dim Action_Plugin_Admin_End()
ReDim Action_Plugin_Admin_End(0)
Dim bAction_Plugin_Admin_End
Dim sAction_Plugin_Admin_End



'**************************************************<
'类型:action
'名称:Action_Plugin_XMLRPC_Begin
'参数:无
'说明:XML-RPC.asp begin
'**************************************************>
Dim Action_Plugin_XMLRPC_Begin()
ReDim Action_Plugin_XMLRPC_Begin(0)
Dim bAction_Plugin_XMLRPC_Begin
Dim sAction_Plugin_XMLRPC_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_XMLRPC_End
'参数:无
'说明:XML-RPC.asp End
'**************************************************>
Dim Action_Plugin_XMLRPC_End()
ReDim Action_Plugin_XMLRPC_End(0)
Dim bAction_Plugin_XMLRPC_End
Dim sAction_Plugin_XMLRPC_End



'**************************************************<
'类型:action
'名称:Action_Plugin_View_Begin
'参数:无
'说明:View.asp Begin
'**************************************************>
Dim Action_Plugin_View_Begin()
ReDim Action_Plugin_View_Begin(0)
Dim bAction_Plugin_View_Begin
Dim sAction_Plugin_View_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_View_End
'参数:无
'说明:View.asp End
'**************************************************>
Dim Action_Plugin_View_End()
ReDim Action_Plugin_View_End(0)
Dim bAction_Plugin_View_End
Dim sAction_Plugin_View_End



'**************************************************<
'类型:action
'名称:Action_Plugin_Guestbook_Begin
'参数:无
'说明:Guestbook.asp
'**************************************************>
Dim Action_Plugin_Guestbook_Begin()
ReDim Action_Plugin_Guestbook_Begin(0)
Dim bAction_Plugin_Guestbook_Begin
Dim sAction_Plugin_Guestbook_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_Guestbook_End
'参数:无
'说明:Guestbook.asp
'**************************************************>
Dim Action_Plugin_Guestbook_End()
ReDim Action_Plugin_Guestbook_End(0)
Dim bAction_Plugin_Guestbook_End
Dim sAction_Plugin_Guestbook_End



'**************************************************<
'类型:action
'名称:Action_Plugin_Feed_Begin
'参数:无
'说明:Feed.asp
'**************************************************>
Dim Action_Plugin_Feed_Begin()
ReDim Action_Plugin_Feed_Begin(0)
Dim bAction_Plugin_Feed_Begin
Dim sAction_Plugin_Feed_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_Feed_End
'参数:无
'说明:Feed.asp
'**************************************************>
Dim Action_Plugin_Feed_End()
ReDim Action_Plugin_Feed_End(0)
Dim bAction_Plugin_Feed_End
Dim sAction_Plugin_Feed_End



'**************************************************<
'类型:action
'名称:Action_Plugin_Wap_Begin
'参数:无
'说明:Wap.asp
'**************************************************>
Dim Action_Plugin_Wap_Begin()
ReDim Action_Plugin_Wap_Begin(0)
Dim bAction_Plugin_Wap_Begin
Dim sAction_Plugin_Wap_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_Wap_End
'参数:无
'说明:Wap.asp
'**************************************************>
Dim Action_Plugin_Wap_End()
ReDim Action_Plugin_Wap_End(0)
Dim bAction_Plugin_Wap_End
Dim sAction_Plugin_Wap_End



'**************************************************<
'类型:action
'名称:Action_Plugin_Catalog_Begin
'参数:无
'说明:Catalog.asp
'**************************************************>
Dim Action_Plugin_Catalog_Begin()
ReDim Action_Plugin_Catalog_Begin(0)
Dim bAction_Plugin_Catalog_Begin
Dim sAction_Plugin_Catalog_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_Catalog_End
'参数:无
'说明:Catalog.asp
'**************************************************>
Dim Action_Plugin_Catalog_End()
ReDim Action_Plugin_Catalog_End(0)
Dim bAction_Plugin_Catalog_End
Dim sAction_Plugin_Catalog_End



'**************************************************<
'类型:action
'名称:Action_Plugin_Searching_Begin
'参数:无
'说明:Search.asp
'**************************************************>
Dim Action_Plugin_Searching_Begin()
ReDim Action_Plugin_Searching_Begin(0)
Dim bAction_Plugin_Searching_Begin
Dim sAction_Plugin_Searching_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_Searching_End
'参数:无
'说明:Search.asp
'**************************************************>
Dim Action_Plugin_Searching_End()
ReDim Action_Plugin_Searching_End(0)
Dim bAction_Plugin_Searching_End
Dim sAction_Plugin_Searching_End



'**************************************************<
'类型:action
'名称:Action_Plugin_Default_Begin
'参数:无
'说明:Default.asp
'**************************************************>
Dim Action_Plugin_Default_Begin()
ReDim Action_Plugin_Default_Begin(0)
Dim bAction_Plugin_Default_Begin
Dim sAction_Plugin_Default_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_Default_End
'参数:无
'说明:Default.asp
'**************************************************>
Dim Action_Plugin_Default_End()
ReDim Action_Plugin_Default_End(0)
Dim bAction_Plugin_Default_End
Dim sAction_Plugin_Default_End



'**************************************************<
'类型:action
'名称:Action_Plugin_Edit_Comment_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_Edit_Comment_Begin()
ReDim Action_Plugin_Edit_Comment_Begin(0)
Dim bAction_Plugin_Edit_Comment_Begin
Dim sAction_Plugin_Edit_Comment_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_Edit_UEditor_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_Edit_UEditor_Begin()
ReDim Action_Plugin_Edit_UEditor_Begin(0)
Dim bAction_Plugin_Edit_UEditor_Begin
Dim sAction_Plugin_Edit_UEditor_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_Edit_Link_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_Edit_Link_Begin()
ReDim Action_Plugin_Edit_Link_Begin(0)
Dim bAction_Plugin_Edit_Link_Begin
Dim sAction_Plugin_Edit_Link_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_Edit_Setting_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_Edit_Setting_Begin()
ReDim Action_Plugin_Edit_Setting_Begin(0)
Dim bAction_Plugin_Edit_Setting_Begin
Dim sAction_Plugin_Edit_Setting_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_Edit_Tag_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_Edit_Tag_Begin()
ReDim Action_Plugin_Edit_Tag_Begin(0)
Dim bAction_Plugin_Edit_Tag_Begin
Dim sAction_Plugin_Edit_Tag_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_Edit_User_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_Edit_User_Begin()
ReDim Action_Plugin_Edit_User_Begin(0)
Dim bAction_Plugin_Edit_User_Begin
Dim sAction_Plugin_Edit_User_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_Edit_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_Edit_Begin()
ReDim Action_Plugin_Edit_Begin(0)
Dim bAction_Plugin_Edit_Begin
Dim sAction_Plugin_Edit_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_Edit_Catalog_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_Edit_Catalog_Begin()
ReDim Action_Plugin_Edit_Catalog_Begin(0)
Dim bAction_Plugin_Edit_Catalog_Begin
Dim sAction_Plugin_Edit_Catalog_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_BlogLogin_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_BlogLogin_Begin()
ReDim Action_Plugin_BlogLogin_Begin(0)
Dim bAction_Plugin_BlogLogin_Begin
Dim sAction_Plugin_BlogLogin_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_BlogVerify_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_BlogVerify_Begin()
ReDim Action_Plugin_BlogVerify_Begin(0)
Dim bAction_Plugin_BlogVerify_Begin
Dim sAction_Plugin_BlogVerify_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_BlogVerify_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_BlogVerify_Succeed()
ReDim Action_Plugin_BlogVerify_Succeed(0)
Dim bAction_Plugin_BlogVerify_Succeed
Dim sAction_Plugin_BlogVerify_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_BlogLogout_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_BlogLogout_Begin()
ReDim Action_Plugin_BlogLogout_Begin(0)
Dim bAction_Plugin_BlogLogout_Begin
Dim sAction_Plugin_BlogLogout_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_BlogLogout_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_BlogLogout_Succeed()
ReDim Action_Plugin_BlogLogout_Succeed(0)
Dim bAction_Plugin_BlogLogout_Succeed
Dim sAction_Plugin_BlogLogout_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_BlogAdmin_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_BlogAdmin_Begin()
ReDim Action_Plugin_BlogAdmin_Begin(0)
Dim bAction_Plugin_BlogAdmin_Begin
Dim sAction_Plugin_BlogAdmin_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_ViewRights_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_ViewRights_Begin()
ReDim Action_Plugin_ViewRights_Begin(0)
Dim bAction_Plugin_ViewRights_Begin
Dim sAction_Plugin_ViewRights_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_ArticleMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_ArticleMng_Begin()
ReDim Action_Plugin_ArticleMng_Begin(0)
Dim bAction_Plugin_ArticleMng_Begin
Dim sAction_Plugin_ArticleMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_ArticleEdt_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_ArticleEdt_Begin()
ReDim Action_Plugin_ArticleEdt_Begin(0)
Dim bAction_Plugin_ArticleEdt_Begin
Dim sAction_Plugin_ArticleEdt_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_ArticlePst_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_ArticlePst_Begin()
ReDim Action_Plugin_ArticlePst_Begin(0)
Dim bAction_Plugin_ArticlePst_Begin
Dim sAction_Plugin_ArticlePst_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_ArticlePst_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_ArticlePst_Succeed()
ReDim Action_Plugin_ArticlePst_Succeed(0)
Dim bAction_Plugin_ArticlePst_Succeed
Dim sAction_Plugin_ArticlePst_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_ArticleDel_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_ArticleDel_Begin()
ReDim Action_Plugin_ArticleDel_Begin(0)
Dim bAction_Plugin_ArticleDel_Begin
Dim sAction_Plugin_ArticleDel_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_ArticleDel_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_ArticleDel_Succeed()
ReDim Action_Plugin_ArticleDel_Succeed(0)
Dim bAction_Plugin_ArticleDel_Succeed
Dim sAction_Plugin_ArticleDel_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_CategoryMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CategoryMng_Begin()
ReDim Action_Plugin_CategoryMng_Begin(0)
Dim bAction_Plugin_CategoryMng_Begin
Dim sAction_Plugin_CategoryMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CategoryEdt_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CategoryEdt_Begin()
ReDim Action_Plugin_CategoryEdt_Begin(0)
Dim bAction_Plugin_CategoryEdt_Begin
Dim sAction_Plugin_CategoryEdt_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CategoryPst_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CategoryPst_Begin()
ReDim Action_Plugin_CategoryPst_Begin(0)
Dim bAction_Plugin_CategoryPst_Begin
Dim sAction_Plugin_CategoryPst_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CategoryDel_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CategoryDel_Begin()
ReDim Action_Plugin_CategoryDel_Begin(0)
Dim bAction_Plugin_CategoryDel_Begin
Dim sAction_Plugin_CategoryDel_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CategoryPst_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CategoryPst_Succeed()
ReDim Action_Plugin_CategoryPst_Succeed(0)
Dim bAction_Plugin_CategoryPst_Succeed
Dim sAction_Plugin_CategoryPst_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_CategoryDel_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CategoryDel_Succeed()
ReDim Action_Plugin_CategoryDel_Succeed(0)
Dim bAction_Plugin_CategoryDel_Succeed
Dim sAction_Plugin_CategoryDel_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentMng_Begin()
ReDim Action_Plugin_CommentMng_Begin(0)
Dim bAction_Plugin_CommentMng_Begin
Dim sAction_Plugin_CommentMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentPost_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentPost_Begin()
ReDim Action_Plugin_CommentPost_Begin(0)
Dim bAction_Plugin_CommentPost_Begin
Dim sAction_Plugin_CommentPost_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentPost_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentPost_Succeed()
ReDim Action_Plugin_CommentPost_Succeed(0)
Dim bAction_Plugin_CommentPost_Succeed
Dim sAction_Plugin_CommentPost_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentDel_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentDel_Begin()
ReDim Action_Plugin_CommentDel_Begin(0)
Dim bAction_Plugin_CommentDel_Begin
Dim sAction_Plugin_CommentDel_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentDel_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentDel_Succeed()
ReDim Action_Plugin_CommentDel_Succeed(0)
Dim bAction_Plugin_CommentDel_Succeed
Dim sAction_Plugin_CommentDel_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentRev_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentRev_Begin()
ReDim Action_Plugin_CommentRev_Begin(0)
Dim bAction_Plugin_CommentRev_Begin
Dim sAction_Plugin_CommentRev_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentRev_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentRev_Succeed()
ReDim Action_Plugin_CommentRev_Succeed(0)
Dim bAction_Plugin_CommentRev_Succeed
Dim sAction_Plugin_CommentRev_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentEdt_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentEdt_Begin()
ReDim Action_Plugin_CommentEdt_Begin(0)
Dim bAction_Plugin_CommentEdt_Begin
Dim sAction_Plugin_CommentEdt_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentSav_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentSav_Begin()
ReDim Action_Plugin_CommentSav_Begin(0)
Dim bAction_Plugin_CommentSav_Begin
Dim sAction_Plugin_CommentSav_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentSav_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentSav_Succeed()
ReDim Action_Plugin_CommentSav_Succeed(0)
Dim bAction_Plugin_CommentSav_Succeed
Dim sAction_Plugin_CommentSav_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_TrackBackMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TrackBackMng_Begin()
ReDim Action_Plugin_TrackBackMng_Begin(0)
Dim bAction_Plugin_TrackBackMng_Begin
Dim sAction_Plugin_TrackBackMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_TrackBackPost_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TrackBackPost_Begin()
ReDim Action_Plugin_TrackBackPost_Begin(0)
Dim bAction_Plugin_TrackBackPost_Begin
Dim sAction_Plugin_TrackBackPost_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_TrackBackPost_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TrackBackPost_Succeed()
ReDim Action_Plugin_TrackBackPost_Succeed(0)
Dim bAction_Plugin_TrackBackPost_Succeed
Dim sAction_Plugin_TrackBackPost_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_TrackBackDel_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TrackBackDel_Begin()
ReDim Action_Plugin_TrackBackDel_Begin(0)
Dim bAction_Plugin_TrackBackDel_Begin
Dim sAction_Plugin_TrackBackDel_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_TrackBackDel_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TrackBackDel_Succeed()
ReDim Action_Plugin_TrackBackDel_Succeed(0)
Dim bAction_Plugin_TrackBackDel_Succeed
Dim sAction_Plugin_TrackBackDel_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_TrackBackSnd_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TrackBackSnd_Begin()
ReDim Action_Plugin_TrackBackSnd_Begin(0)
Dim bAction_Plugin_TrackBackSnd_Begin
Dim sAction_Plugin_TrackBackSnd_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_TrackBackSnd_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TrackBackSnd_Succeed()
ReDim Action_Plugin_TrackBackSnd_Succeed(0)
Dim bAction_Plugin_TrackBackSnd_Succeed
Dim sAction_Plugin_TrackBackSnd_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_UserMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_UserMng_Begin()
ReDim Action_Plugin_UserMng_Begin(0)
Dim bAction_Plugin_UserMng_Begin
Dim sAction_Plugin_UserMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_UserCrt_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_UserCrt_Begin()
ReDim Action_Plugin_UserCrt_Begin(0)
Dim bAction_Plugin_UserCrt_Begin
Dim sAction_Plugin_UserCrt_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_UserEdt_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_UserEdt_Begin()
ReDim Action_Plugin_UserEdt_Begin(0)
Dim bAction_Plugin_UserEdt_Begin
Dim sAction_Plugin_UserEdt_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_UserDel_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_UserDel_Begin()
ReDim Action_Plugin_UserDel_Begin(0)
Dim bAction_Plugin_UserDel_Begin
Dim sAction_Plugin_UserDel_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_UserCrt_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_UserCrt_Succeed()
ReDim Action_Plugin_UserCrt_Succeed(0)
Dim bAction_Plugin_UserCrt_Succeed
Dim sAction_Plugin_UserCrt_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_UserEdt_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_UserEdt_Succeed()
ReDim Action_Plugin_UserEdt_Succeed(0)
Dim bAction_Plugin_UserEdt_Succeed
Dim sAction_Plugin_UserEdt_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_UserDel_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_UserDel_Succeed()
ReDim Action_Plugin_UserDel_Succeed(0)
Dim bAction_Plugin_UserDel_Succeed
Dim sAction_Plugin_UserDel_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_FileMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_FileMng_Begin()
ReDim Action_Plugin_FileMng_Begin(0)
Dim bAction_Plugin_FileMng_Begin
Dim sAction_Plugin_FileMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_FileSnd_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_FileSnd_Begin()
ReDim Action_Plugin_FileSnd_Begin(0)
Dim bAction_Plugin_FileSnd_Begin
Dim sAction_Plugin_FileSnd_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_FileUpload_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_FileUpload_Begin()
ReDim Action_Plugin_FileUpload_Begin(0)
Dim bAction_Plugin_FileUpload_Begin
Dim sAction_Plugin_FileUpload_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_FileDel_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_FileDel_Begin()
ReDim Action_Plugin_FileDel_Begin(0)
Dim bAction_Plugin_FileDel_Begin
Dim sAction_Plugin_FileDel_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_FileUpload_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_FileUpload_Succeed()
ReDim Action_Plugin_FileUpload_Succeed(0)
Dim bAction_Plugin_FileUpload_Succeed
Dim sAction_Plugin_FileUpload_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_FileDel_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_FileDel_Succeed()
ReDim Action_Plugin_FileDel_Succeed(0)
Dim bAction_Plugin_FileDel_Succeed
Dim sAction_Plugin_FileDel_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_Search_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_Search_Begin()
ReDim Action_Plugin_Search_Begin(0)
Dim bAction_Plugin_Search_Begin
Dim sAction_Plugin_Search_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_SettingMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_SettingMng_Begin()
ReDim Action_Plugin_SettingMng_Begin(0)
Dim bAction_Plugin_SettingMng_Begin
Dim sAction_Plugin_SettingMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_SettingSav_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_SettingSav_Begin()
ReDim Action_Plugin_SettingSav_Begin(0)
Dim bAction_Plugin_SettingSav_Begin
Dim sAction_Plugin_SettingSav_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_SettingSav_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_SettingSav_Succeed()
ReDim Action_Plugin_SettingSav_Succeed(0)
Dim bAction_Plugin_SettingSav_Succeed
Dim sAction_Plugin_SettingSav_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_TagMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TagMng_Begin()
ReDim Action_Plugin_TagMng_Begin(0)
Dim bAction_Plugin_TagMng_Begin
Dim sAction_Plugin_TagMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_TagEdt_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TagEdt_Begin()
ReDim Action_Plugin_TagEdt_Begin(0)
Dim bAction_Plugin_TagEdt_Begin
Dim sAction_Plugin_TagEdt_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_TagPst_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TagPst_Begin()
ReDim Action_Plugin_TagPst_Begin(0)
Dim bAction_Plugin_TagPst_Begin
Dim sAction_Plugin_TagPst_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_TagDel_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TagDel_Begin()
ReDim Action_Plugin_TagDel_Begin(0)
Dim bAction_Plugin_TagDel_Begin
Dim sAction_Plugin_TagDel_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_TagPst_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TagPst_Succeed()
ReDim Action_Plugin_TagPst_Succeed(0)
Dim bAction_Plugin_TagPst_Succeed
Dim sAction_Plugin_TagPst_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_TagDel_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TagDel_Succeed()
ReDim Action_Plugin_TagDel_Succeed(0)
Dim bAction_Plugin_TagDel_Succeed
Dim sAction_Plugin_TagDel_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_BlogReBuild_Begin()
ReDim Action_Plugin_BlogReBuild_Begin(0)
Dim bAction_Plugin_BlogReBuild_Begin
Dim sAction_Plugin_BlogReBuild_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_BlogReBuild_Succeed()
ReDim Action_Plugin_BlogReBuild_Succeed(0)
Dim bAction_Plugin_BlogReBuild_Succeed
Dim sAction_Plugin_BlogReBuild_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_DirectoryReBuild_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_DirectoryReBuild_Begin()
ReDim Action_Plugin_DirectoryReBuild_Begin(0)
Dim bAction_Plugin_DirectoryReBuild_Begin
Dim sAction_Plugin_DirectoryReBuild_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_DirectoryReBuild_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_DirectoryReBuild_Succeed()
ReDim Action_Plugin_DirectoryReBuild_Succeed(0)
Dim bAction_Plugin_DirectoryReBuild_Succeed
Dim sAction_Plugin_DirectoryReBuild_Succeed



'**************************************************<
'类型:action
'名称:Action_Plugin_FileReBuild_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_FileReBuild_Begin()
ReDim Action_Plugin_FileReBuild_Begin(0)
Dim bAction_Plugin_FileReBuild_Begin
Dim sAction_Plugin_FileReBuild_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_FileReBuild_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_FileReBuild_Succeed()
ReDim Action_Plugin_FileReBuild_Succeed(0)
Dim bAction_Plugin_FileReBuild_Succeed
Dim sAction_Plugin_FileReBuild_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_SiteInfo_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_SiteInfo_Begin()
ReDim Action_Plugin_SiteInfo_Begin(0)
Dim bAction_Plugin_SiteInfo_Begin
Dim sAction_Plugin_SiteInfo_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_SiteFileMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_SiteFileMng_Begin()
ReDim Action_Plugin_SiteFileMng_Begin(0)
Dim bAction_Plugin_SiteFileMng_Begin
Dim sAction_Plugin_SiteFileMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_SiteFileEdt_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_SiteFileEdt_Begin()
ReDim Action_Plugin_SiteFileEdt_Begin(0)
Dim bAction_Plugin_SiteFileEdt_Begin
Dim sAction_Plugin_SiteFileEdt_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_SiteFilePst_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_SiteFilePst_Begin()
ReDim Action_Plugin_SiteFilePst_Begin(0)
Dim bAction_Plugin_SiteFilePst_Begin
Dim sAction_Plugin_SiteFilePst_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_SiteFileDel_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_SiteFileDel_Begin()
ReDim Action_Plugin_SiteFileDel_Begin(0)
Dim bAction_Plugin_SiteFileDel_Begin
Dim sAction_Plugin_SiteFileDel_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_SiteFilePst_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_SiteFilePst_Succeed()
ReDim Action_Plugin_SiteFilePst_Succeed(0)
Dim bAction_Plugin_SiteFilePst_Succeed
Dim sAction_Plugin_SiteFilePst_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_SiteFileDel_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_SiteFileDel_Succeed()
ReDim Action_Plugin_SiteFileDel_Succeed(0)
Dim bAction_Plugin_SiteFileDel_Succeed
Dim sAction_Plugin_SiteFileDel_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_AskFileReBuild_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_AskFileReBuild_Begin()
ReDim Action_Plugin_AskFileReBuild_Begin(0)
Dim bAction_Plugin_AskFileReBuild_Begin
Dim sAction_Plugin_AskFileReBuild_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_TrackBackUrlGet_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TrackBackUrlGet_Begin()
ReDim Action_Plugin_TrackBackUrlGet_Begin(0)
Dim bAction_Plugin_TrackBackUrlGet_Begin
Dim sAction_Plugin_TrackBackUrlGet_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentDelBatch_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentDelBatch_Begin()
ReDim Action_Plugin_CommentDelBatch_Begin(0)
Dim bAction_Plugin_CommentDelBatch_Begin
Dim sAction_Plugin_CommentDelBatch_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_TrackBackDelBatch_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TrackBackDelBatch_Begin()
ReDim Action_Plugin_TrackBackDelBatch_Begin(0)
Dim bAction_Plugin_TrackBackDelBatch_Begin
Dim sAction_Plugin_TrackBackDelBatch_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_FileDelBatch_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_FileDelBatch_Begin()
ReDim Action_Plugin_FileDelBatch_Begin(0)
Dim bAction_Plugin_FileDelBatch_Begin
Dim sAction_Plugin_FileDelBatch_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_CommentDelBatch_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_CommentDelBatch_Succeed()
ReDim Action_Plugin_CommentDelBatch_Succeed(0)
Dim bAction_Plugin_CommentDelBatch_Succeed
Dim sAction_Plugin_CommentDelBatch_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_TrackBackDelBatch_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TrackBackDelBatch_Succeed()
ReDim Action_Plugin_TrackBackDelBatch_Succeed(0)
Dim bAction_Plugin_TrackBackDelBatch_Succeed
Dim sAction_Plugin_TrackBackDelBatch_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_FileDelBatch_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_FileDelBatch_Succeed()
ReDim Action_Plugin_FileDelBatch_Succeed(0)
Dim bAction_Plugin_FileDelBatch_Succeed
Dim sAction_Plugin_FileDelBatch_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_LinkMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_LinkMng_Begin()
ReDim Action_Plugin_LinkMng_Begin(0)
Dim bAction_Plugin_LinkMng_Begin
Dim sAction_Plugin_LinkMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_LinkSav_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_LinkSav_Begin()
ReDim Action_Plugin_LinkSav_Begin(0)
Dim bAction_Plugin_LinkSav_Begin
Dim sAction_Plugin_LinkSav_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_LinkSav_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_LinkSav_Succeed()
ReDim Action_Plugin_LinkSav_Succeed(0)
Dim bAction_Plugin_LinkSav_Succeed
Dim sAction_Plugin_LinkSav_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_PlugInMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_PlugInMng_Begin()
ReDim Action_Plugin_PlugInMng_Begin(0)
Dim bAction_Plugin_PlugInMng_Begin
Dim sAction_Plugin_PlugInMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_PlugInActive_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_PlugInActive_Begin()
ReDim Action_Plugin_PlugInActive_Begin(0)
Dim bAction_Plugin_PlugInActive_Begin
Dim sAction_Plugin_PlugInActive_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_PlugInDisable_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_PlugInDisable_Begin()
ReDim Action_Plugin_PlugInDisable_Begin(0)
Dim bAction_Plugin_PlugInDisable_Begin
Dim sAction_Plugin_PlugInDisable_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_PlugInActive_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_PlugInActive_Succeed()
ReDim Action_Plugin_PlugInActive_Succeed(0)
Dim bAction_Plugin_PlugInActive_Succeed
Dim sAction_Plugin_PlugInActive_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_PlugInDisable_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_PlugInDisable_Succeed()
ReDim Action_Plugin_PlugInDisable_Succeed(0)
Dim bAction_Plugin_PlugInDisable_Succeed
Dim sAction_Plugin_PlugInDisable_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_ThemeMng_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_ThemeMng_Begin()
ReDim Action_Plugin_ThemeMng_Begin(0)
Dim bAction_Plugin_ThemeMng_Begin
Dim sAction_Plugin_ThemeMng_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_ThemeSav_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_ThemeSav_Begin()
ReDim Action_Plugin_ThemeSav_Begin(0)
Dim bAction_Plugin_ThemeSav_Begin
Dim sAction_Plugin_ThemeSav_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_ThemeSav_Succeed
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_Themesav_Succeed()
ReDim Action_Plugin_ThemeSav_Succeed(0)
Dim bAction_Plugin_ThemeSav_Succeed
Dim sAction_Plugin_ThemeSav_Succeed


'**************************************************<
'类型:action
'名称:Action_Plugin_MakeBlogReBuild_Begin
'参数:无
'说明:执行重建索引
'**************************************************>
Dim Action_Plugin_MakeBlogReBuild_Begin()
ReDim Action_Plugin_MakeBlogReBuild_Begin(0)
Dim bAction_Plugin_MakeBlogReBuild_Begin
Dim sAction_Plugin_MakeBlogReBuild_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_MakeBlogReBuild_End
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_MakeBlogReBuild_End()
ReDim Action_Plugin_MakeBlogReBuild_End(0)
Dim bAction_Plugin_MakeBlogReBuild_End
Dim sAction_Plugin_MakeBlogReBuild_End


'**************************************************<
'类型:action
'名称:Action_Plugin_MakeBlogReBuild_Core_Begin
'参数:无
'说明:执行重建索引的核心过程
'**************************************************>
Dim Action_Plugin_MakeBlogReBuild_Core_Begin()
ReDim Action_Plugin_MakeBlogReBuild_Core_Begin(0)
Dim bAction_Plugin_MakeBlogReBuild_Core_Begin
Dim sAction_Plugin_MakeBlogReBuild_Core_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_MakeBlogReBuild_Core_End
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_MakeBlogReBuild_Core_End()
ReDim Action_Plugin_MakeBlogReBuild_Core_End(0)
Dim bAction_Plugin_MakeBlogReBuild_Core_End
Dim sAction_Plugin_MakeBlogReBuild_Core_End


'**************************************************<
'类型:action
'名称:Action_Plugin_MakeFileReBuild_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_MakeFileReBuild_Begin()
ReDim Action_Plugin_MakeFileReBuild_Begin(0)
Dim bAction_Plugin_MakeFileReBuild_Begin
Dim sAction_Plugin_MakeFileReBuild_Begin


'**************************************************<
'类型:action
'名称:Action_Plugin_MakeFileReBuild_End
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_MakeFileReBuild_End()
ReDim Action_Plugin_MakeFileReBuild_End(0)
Dim bAction_Plugin_MakeFileReBuild_End
Dim sAction_Plugin_MakeFileReBuild_End



'**************************************************<
'类型:action
'名称:Action_Plugin_MakeCalendar_Begin
'参数:无
'说明:执行日历生成
'**************************************************>
Dim Action_Plugin_MakeCalendar_Begin()
ReDim Action_Plugin_MakeCalendar_Begin(0)
Dim bAction_Plugin_MakeCalendar_Begin
Dim sAction_Plugin_MakeCalendar_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_GetRights_Begin
'参数:无
'说明:GetRights
'**************************************************>
Dim Action_Plugin_GetRights_Begin()
ReDim Action_Plugin_GetRights_Begin(0)
Dim bAction_Plugin_GetRights_Begin
Dim sAction_Plugin_GetRights_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_CheckRights_Begin
'参数:无
'说明:CheckRights
'**************************************************>
Dim Action_Plugin_CheckRights_Begin()
ReDim Action_Plugin_CheckRights_Begin(0)
Dim bAction_Plugin_CheckRights_Begin
Dim sAction_Plugin_CheckRights_Begin




'**************************************************<
'类型:action
'名称:Action_Plugin_TArticle_Export_Begin
'参数:无
'说明:TArticle_Export_Begin
'**************************************************>
Dim Action_Plugin_TArticle_Export_Begin()
ReDim Action_Plugin_TArticle_Export_Begin(0)
Dim bAction_Plugin_TArticle_Export_Begin
Dim sAction_Plugin_TArticle_Export_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticle_Export_End
'参数:无
'说明:TArticle_Export_End
'**************************************************>
Dim Action_Plugin_TArticle_Export_End()
ReDim Action_Plugin_TArticle_Export_End(0)
Dim bAction_Plugin_TArticle_Export_End
Dim sAction_Plugin_TArticle_Export_End



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticle_Export_Tag_Begin
'参数:无
'说明:TArticle_Export_Tag_Begin
'**************************************************>
Dim Action_Plugin_TArticle_Export_Tag_Begin()
ReDim Action_Plugin_TArticle_Export_Tag_Begin(0)
Dim bAction_Plugin_TArticle_Export_Tag_Begin
Dim sAction_Plugin_TArticle_Export_Tag_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticle_Export_CMTandTB_Begin
'参数:无
'说明:TArticle_Export_CMTandTB_Begin
'**************************************************>
Dim Action_Plugin_TArticle_Export_CMTandTB_Begin()
ReDim Action_Plugin_TArticle_Export_CMTandTB_Begin(0)
Dim bAction_Plugin_TArticle_Export_CMTandTB_Begin
Dim sAction_Plugin_TArticle_Export_CMTandTB_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticle_Export_NavBar_Begin
'参数:无
'说明:TArticle_Export_NavBar_Begin
'**************************************************>
Dim Action_Plugin_TArticle_Export_NavBar_Begin()
ReDim Action_Plugin_TArticle_Export_NavBar_Begin(0)
Dim bAction_Plugin_TArticle_Export_NavBar_Begin
Dim sAction_Plugin_TArticle_Export_NavBar_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticle_Export_CommentPost_Begin
'参数:无
'说明:TArticle_Export_CommentPost_Begin
'**************************************************>
Dim Action_Plugin_TArticle_Export_CommentPost_Begin()
ReDim Action_Plugin_TArticle_Export_CommentPost_Begin(0)
Dim bAction_Plugin_TArticle_Export_CommentPost_Begin
Dim sAction_Plugin_TArticle_Export_CommentPost_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticle_Export_Mutuality_Begin
'参数:无
'说明:TArticle_Export_Mutuality_Begin
'**************************************************>
Dim Action_Plugin_TArticle_Export_Mutuality_Begin()
ReDim Action_Plugin_TArticle_Export_Mutuality_Begin(0)
Dim bAction_Plugin_TArticle_Export_Mutuality_Begin
Dim sAction_Plugin_TArticle_Export_Mutuality_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticleList_Export_Begin
'参数:无
'说明:TArticleList_Export_Begin
'**************************************************>
Dim Action_Plugin_TArticleList_Export_Begin()
ReDim Action_Plugin_TArticleList_Export_Begin(0)
Dim bAction_Plugin_TArticleList_Export_Begin
Dim sAction_Plugin_TArticleList_Export_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticleList_Export_End
'参数:无
'说明:TArticleList_Export_End
'**************************************************>
Dim Action_Plugin_TArticleList_Export_End()
ReDim Action_Plugin_TArticleList_Export_End(0)
Dim bAction_Plugin_TArticleList_Export_End
Dim sAction_Plugin_TArticleList_Export_End



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticleList_ExportByCache_Begin
'参数:无
'说明:TArticleList_ExportByCache_Begin
'**************************************************>
Dim Action_Plugin_TArticleList_ExportByCache_Begin()
ReDim Action_Plugin_TArticleList_ExportByCache_Begin(0)
Dim bAction_Plugin_TArticleList_ExportByCache_Begin
Dim sAction_Plugin_TArticleList_ExportByCache_Begin



'**************************************************<
'类型:action
'名称:TArticleList_ExportByMixed_Begin
'参数:无
'说明:
'**************************************************>
Dim Action_Plugin_TArticleList_ExportByMixed_Begin()
ReDim Action_Plugin_TArticleList_ExportByMixed_Begin(0)
Dim bAction_Plugin_TArticleList_ExportByMixed_Begin
Dim sAction_Plugin_TArticleList_ExportByMixed_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticleList_ExportByCache_End
'参数:无
'说明:TArticleList_ExportByCache_End
'**************************************************>
Dim Action_Plugin_TArticleList_ExportByCache_End()
ReDim Action_Plugin_TArticleList_ExportByCache_End(0)
Dim bAction_Plugin_TArticleList_ExportByCache_End
Dim sAction_Plugin_TArticleList_ExportByCache_End



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticleList_ExportByMixed_End
'参数:无
'说明:TArticleList_ExportByMixed_End
'**************************************************>
Dim Action_Plugin_TArticleList_ExportByMixed_End()
ReDim Action_Plugin_TArticleList_ExportByMixed_End(0)
Dim bAction_Plugin_TArticleList_ExportByMixed_End
Dim sAction_Plugin_TArticleList_ExportByMixed_End



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticleList_ExportBar_Begin
'参数:无
'说明:TArticleList_ExportBar_Begin
'**************************************************>
Dim Action_Plugin_TArticleList_ExportBar_Begin()
ReDim Action_Plugin_TArticleList_ExportBar_Begin(0)
Dim bAction_Plugin_TArticleList_ExportBar_Begin
Dim sAction_Plugin_TArticleList_ExportBar_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticleList_ExportBar_End
'参数:无
'说明:TArticleList_ExportBar_End
'**************************************************>
Dim Action_Plugin_TArticleList_ExportBar_End()
ReDim Action_Plugin_TArticleList_ExportBar_End(0)
Dim bAction_Plugin_TArticleList_ExportBar_End
Dim sAction_Plugin_TArticleList_ExportBar_End


'
'**************************************************<
'类型:action
'名称:Action_Plugin_TArticleList_Search_Begin
'参数:无
'说明:TArticleList_Search_Begin
'**************************************************>
Dim Action_Plugin_TArticleList_Search_Begin()
ReDim Action_Plugin_TArticleList_Search_Begin(0)
Dim bAction_Plugin_TArticleList_Search_Begin
Dim sAction_Plugin_TArticleList_Search_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TArticleList_Search_End
'参数:无
'说明:TArticleList_Search_End
'**************************************************>
Dim Action_Plugin_TArticleList_Search_End()
ReDim Action_Plugin_TArticleList_Search_End(0)
Dim bAction_Plugin_TArticleList_Search_End
Dim sAction_Plugin_TArticleList_Search_End






'***************
'2.filter
'***************




'**************************************************<
'类型:filter
'名称:Filter_Plugin_PostComment_Core
'参数:objComment
'说明:发表评论接口
'调用:c_system_event的PostComment,RevertComment
'**************************************************>
Dim sFilter_Plugin_PostComment_Core
Function Filter_Plugin_PostComment_Core(ByRef objComment)
	Dim s,i

	If sFilter_Plugin_PostComment_Core="" Then Exit Function

	s=Split(sFilter_Plugin_PostComment_Core,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objComment")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_PostTrackBack_Core
'参数:objTrackBack
'说明:接收引用接口
'调用:c_system_event的PostTrackBack
'**************************************************>
Dim sFilter_Plugin_PostTrackBack_Core
Function Filter_Plugin_PostTrackBack_Core(ByRef objTrackBack)

	Dim s,i

	If sFilter_Plugin_PostTrackBack_Core="" Then Exit Function

	s=Split(sFilter_Plugin_PostTrackBack_Core,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objTrackBack")
	Next

End Function




'**************************************************<
'类型:filter
'名称:Filter_Plugin_PostTrackBack_Succeed
'参数:objTrackBack
'说明:接收引用接口
'调用:c_system_event的PostTrackBack
'**************************************************>
Dim sFilter_Plugin_PostTrackBack_Succeed
Function Filter_Plugin_PostTrackBack_Succeed(ByRef objTrackBack)

	Dim s,i

	If sFilter_Plugin_PostTrackBack_Succeed="" Then Exit Function

	s=Split(sFilter_Plugin_PostTrackBack_Succeed,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objTrackBack")
	Next

End Function




'**************************************************<
'类型:filter
'名称:Filter_Plugin_PostArticle_Core
'参数:objArticle
'说明:PostArticle
'调用:c_system_event的PostArticle
'**************************************************>
Dim sFilter_Plugin_PostArticle_Core
Function Filter_Plugin_PostArticle_Core(ByRef objArticle)

	Dim s,i

	If sFilter_Plugin_PostArticle_Core="" Then Exit Function

	s=Split(sFilter_Plugin_PostArticle_Core,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objArticle")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_PostArticle_Succeed
'参数:objArticle
'说明:PostArticle
'调用:c_system_event的PostArticle
'**************************************************>
Dim sFilter_Plugin_PostArticle_Succeed
Function Filter_Plugin_PostArticle_Succeed(ByRef objArticle)

	Dim s,i

	If sFilter_Plugin_PostArticle_Succeed="" Then Exit Function

	s=Split(sFilter_Plugin_PostArticle_Succeed,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objArticle")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_PostCategory_Core
'参数:objCategory
'说明:发表Category接口
'调用:c_system_event的PostCategory
'**************************************************>
Dim sFilter_Plugin_PostCategory_Core
Function Filter_Plugin_PostCategory_Core(ByRef objCategory)

	Dim s,i

	If sFilter_Plugin_PostCategory_Core="" Then Exit Function

	s=Split(sFilter_Plugin_PostCategory_Core,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objCategory")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_PostCategory_Succeed
'参数:objCategory
'说明:发表Category接口
'调用:c_system_event的PostCategory
'**************************************************>
Dim sFilter_Plugin_PostCategory_Succeed
Function Filter_Plugin_PostCategory_Succeed(ByRef objCategory)

	Dim s,i

	If sFilter_Plugin_PostCategory_Succeed="" Then Exit Function

	s=Split(sFilter_Plugin_PostCategory_Succeed,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objCategory")
	Next

End Function




'**************************************************<
'类型:filter
'名称:Filter_Plugin_PostComment_Succeed
'参数:objComment
'说明:发表Comment接口
'调用:c_system_event的PostComment
'**************************************************>
Dim sFilter_Plugin_PostComment_Succeed
Function Filter_Plugin_PostComment_Succeed(ByRef objComment)

	Dim s,i

	If sFilter_Plugin_PostComment_Succeed="" Then Exit Function

	s=Split(sFilter_Plugin_PostComment_Succeed,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objComment")
	Next

End Function




'**************************************************<
'类型:filter
'名称:Filter_Plugin_EditUser_Core
'参数:objUser
'说明:EditUser接口
'调用:c_system_event的EditUser
'**************************************************>
Dim sFilter_Plugin_EditUser_Core
Function Filter_Plugin_EditUser_Core(ByRef objUser)

	Dim s,i

	If sFilter_Plugin_EditUser_Core="" Then Exit Function

	s=Split(sFilter_Plugin_EditUser_Core,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objUser")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_EditUser_Succeed
'参数:objUser
'说明:EditUser接口
'调用:c_system_event的EditUser
'**************************************************>
Dim sFilter_Plugin_EditUser_Succeed
Function Filter_Plugin_EditUser_Succeed(ByRef objUser)

	Dim s,i

	If sFilter_Plugin_EditUser_Succeed="" Then Exit Function

	s=Split(sFilter_Plugin_EditUser_Succeed,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objUser")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_PostTag_Core
'参数:objTag
'说明:PostTag
'调用:c_system_event的PostTag
'**************************************************>
Dim sFilter_Plugin_PostTag_Core
Function Filter_Plugin_PostTag_Core(ByRef objTag)

	Dim s,i

	If sFilter_Plugin_PostTag_Core="" Then Exit Function

	s=Split(sFilter_Plugin_PostTag_Core,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objTag")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_PostTag_Succeed
'参数:objTag
'说明:PostTag
'调用:c_system_event的PostTag
'**************************************************>
Dim sFilter_Plugin_PostTag_Succeed
Function Filter_Plugin_PostTag_Succeed(ByRef objTag)

	Dim s,i

	If sFilter_Plugin_PostTag_Succeed="" Then Exit Function

	s=Split(sFilter_Plugin_PostTag_Succeed,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "objTag")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticle_Export_Template
'参数:html,Template_Article_Single,Template_Article_Multi,Template_Article_Istop
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticle_Export_Template
Function Filter_Plugin_TArticle_Export_Template(ByRef html,ByRef Template_Article_Single,ByRef Template_Article_Multi,ByRef Template_Article_Istop)

	Dim s,i

	If sFilter_Plugin_TArticle_Export_Template="" Then Exit Function

	s=Split(sFilter_Plugin_TArticle_Export_Template,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " "& "html,Template_Article_Single,Template_Article_Multi,Template_Article_Istop")
	Next

End Function




'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticle_Export_Template_Sub
'参数:Template_Article_Comment,Template_Article_Trackback,Template_Article_Tag,Template_Article_Commentpost,Template_Article_Navbar_L,Template_Article_Navbar_R,Template_Article_Mutuality
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticle_Export_Template_Sub
Function Filter_Plugin_TArticle_Export_Template_Sub(ByRef Template_Article_Comment,ByRef Template_Article_Trackback,ByRef Template_Article_Tag,ByRef Template_Article_Commentpost,ByRef Template_Article_Navbar_L,ByRef Template_Article_Navbar_R,ByRef Template_Article_Mutuality)

	Dim s,i

	If sFilter_Plugin_TArticle_Export_Template_Sub="" Then Exit Function

	s=Split(sFilter_Plugin_TArticle_Export_Template_Sub,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "Template_Article_Comment,Template_Article_Trackback,Template_Article_Tag,Template_Article_Commentpost,Template_Article_Navbar_L,Template_Article_Navbar_R,Template_Article_Mutuality")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticle_Export_TemplateTags
'参数:aryTemplateTagsName,aryTemplateTagsValue
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticle_Export_TemplateTags
Function Filter_Plugin_TArticle_Export_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)

	Dim s,i

	If sFilter_Plugin_TArticle_Export_TemplateTags="" Then Exit Function

	s=Split(sFilter_Plugin_TArticle_Export_TemplateTags,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "aryTemplateTagsName,aryTemplateTagsValue")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticle_Build_TemplateTags
'参数:aryTemplateTagsName,aryTemplateTagsValue
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticle_Build_TemplateTags
Function Filter_Plugin_TArticle_Build_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)

	Dim s,i

	If sFilter_Plugin_TArticle_Build_TemplateTags="" Then Exit Function

	s=Split(sFilter_Plugin_TArticle_Build_TemplateTags,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "aryTemplateTagsName,aryTemplateTagsValue")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticle_Build_Template
'参数:html,wapHtml
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticle_Build_Template
Function Filter_Plugin_TArticle_Build_Template(ByRef html)

	Dim s,i

	If sFilter_Plugin_TArticle_Build_Template="" Then Exit Function

	s=Split(sFilter_Plugin_TArticle_Build_Template,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "html")
	Next

End Function






'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticleList_Export
'参数:intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticleList_Export
Function Filter_Plugin_TArticleList_Export(ByRef intPage,ByRef intCateId,ByRef intAuthorId,ByRef dtmYearMonth,ByRef strTagsName,ByRef intType)

	Dim s,i

	If sFilter_Plugin_TArticleList_Export="" Then Exit Function

	s=Split(sFilter_Plugin_TArticleList_Export,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticleList_ExportByCache
'参数:intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticleList_ExportByCache
Function Filter_Plugin_TArticleList_ExportByCache(ByRef intPage,ByRef intCateId,ByRef intAuthorId,ByRef dtmYearMonth,ByRef strTagsName,ByRef intType)

	Dim s,i

	If sFilter_Plugin_TArticleList_ExportByCache="" Then Exit Function

	s=Split(sFilter_Plugin_TArticleList_ExportByCache,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticleList_ExportByMixed
'参数:intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticleList_ExportByMixed
Function Filter_Plugin_TArticleList_ExportByMixed(ByRef intPage,ByRef intCateId,ByRef intAuthorId,ByRef dtmYearMonth,ByRef strTagsName,ByRef intType)

	Dim s,i

	If sFilter_Plugin_TArticleList_ExportByMixed="" Then Exit Function

	s=Split(sFilter_Plugin_TArticleList_ExportByMixed,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "intPage,intCateId,intAuthorId,dtmYearMonth,strTagsName,intType")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticleList_Build_Template
'参数:html
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticleList_Build_Template
Function Filter_Plugin_TArticleList_Build_Template(ByRef html)

	Dim s,i

	If sFilter_Plugin_TArticleList_Build_Template="" Then Exit Function

	s=Split(sFilter_Plugin_TArticleList_Build_Template,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "html")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticleList_Build_TemplateSub
'参数:aryTemplateSubName,aryTemplateSubValue
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticleList_Build_TemplateSub
Function Filter_Plugin_TArticleList_Build_TemplateSub(ByRef aryTemplateSubName,ByRef aryTemplateSubValue)

	Dim s,i

	If sFilter_Plugin_TArticleList_Build_TemplateSub="" Then Exit Function

	s=Split(sFilter_Plugin_TArticleList_Build_TemplateSub,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "aryTemplateSubName,aryTemplateSubValue")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticleList_Build_TemplateTags
'参数:aryTemplateTagsName,aryTemplateTagsValue
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticleList_Build_TemplateTags
Function Filter_Plugin_TArticleList_Build_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)

	Dim s,i

	If sFilter_Plugin_TArticleList_Build_TemplateTags="" Then Exit Function

	s=Split(sFilter_Plugin_TArticleList_Build_TemplateTags,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "aryTemplateTagsName,aryTemplateTagsValue")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TCategory_Post
'参数:ID,Name,Alias,Order,Count
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TCategory_Post
Function Filter_Plugin_TCategory_Post(ByRef ID,ByRef Name,ByRef Intro,ByRef Order,ByRef Count,ByRef ParentID,ByRef Alias,ByRef TemplateName,ByRef FullUrl,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TCategory_Post="" Then Exit Function

	s=Split(sFilter_Plugin_TCategory_Post,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TCategory_LoadInfoByID
'参数:ID,Name,Alias,Order,Count
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TCategory_LoadInfoByID
Function Filter_Plugin_TCategory_LoadInfoByID(ByRef ID,ByRef Name,ByRef Intro,ByRef Order,ByRef Count,ByRef ParentID,ByRef Alias,ByRef TemplateName,ByRef FullUrl,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TCategory_LoadInfoByID="" Then Exit Function

	s=Split(sFilter_Plugin_TCategory_LoadInfoByID,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TCategory_LoadInfoByArray
'参数:ID,Name,Alias,Order,Count
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TCategory_LoadInfoByArray
Function Filter_Plugin_TCategory_LoadInfoByArray(ByRef ID,ByRef Name,ByRef Intro,ByRef Order,ByRef Count,ByRef ParentID,ByRef Alias,ByRef TemplateName,ByRef FullUrl,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TCategory_LoadInfoByArray="" Then Exit Function

	s=Split(sFilter_Plugin_TCategory_LoadInfoByArray,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString")
	Next

End Function




'**************************************************<
'类型:filter
'名称:Filter_Plugin_TCategory_Del
'参数:ID,Name,Alias,Order,Count
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TCategory_Del
Function Filter_Plugin_TCategory_Del(ByRef ID,ByRef Name,ByRef Intro,ByRef Order,ByRef Count,ByRef ParentID,ByRef Alias,ByRef TemplateName,ByRef FullUrl,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TCategory_Del="" Then Exit Function

	s=Split(sFilter_Plugin_TCategory_Del,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticle_LoadInfobyID
'参数:ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticle_LoadInfobyID
Function Filter_Plugin_TArticle_LoadInfobyID(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef IsAnonymous,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TArticle_LoadInfobyID="" Then Exit Function

	s=Split(sFilter_Plugin_TArticle_LoadInfobyID,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,FullUrl,IsAnonymous,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticle_LoadInfoByArray
'参数:ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticle_LoadInfoByArray
Function Filter_Plugin_TArticle_LoadInfoByArray(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef IsAnonymous,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TArticle_LoadInfoByArray="" Then Exit Function

	s=Split(sFilter_Plugin_TArticle_LoadInfoByArray,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,FullUrl,IsAnonymous,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticle_Del
'参数:ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticle_Del
Function Filter_Plugin_TArticle_Del(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef IsAnonymous,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TArticle_Del="" Then Exit Function

	s=Split(sFilter_Plugin_TArticle_Del,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,FullUrl,IsAnonymous,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticle_Post
'参数:ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticle_Post
Function Filter_Plugin_TArticle_Post(ByRef ID,ByRef Tag,ByRef CateID,ByRef Title,ByRef Intro,ByRef Content,ByRef Level,ByRef AuthorID,ByRef PostTime,ByRef CommNums,ByRef ViewNums,ByRef TrackBackNums,ByRef Alias,ByRef Istop,ByRef TemplateName,ByRef FullUrl,ByRef IsAnonymous,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TArticle_Post="" Then Exit Function

	s=Split(sFilter_Plugin_TArticle_Post,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Tag,CateID,Title,Intro,Content,Level,AuthorID,PostTime,CommNums,ViewNums,TrackBackNums,Alias,Istop,TemplateName,FullUrl,IsAnonymous,MetaString")
	Next

End Function




'**************************************************<
'类型:filter
'名称:Filter_Plugin_TUser_LoadInfobyID
'参数:ID,Name,Level,Password,Email,HomePage,Count,Alias
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TUser_LoadInfobyID
Function Filter_Plugin_TUser_LoadInfobyID(ByRef ID,ByRef Name,ByRef Level,ByRef Password,ByRef Email,ByRef HomePage,ByRef Count,ByRef Alias,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TUser_LoadInfobyID="" Then Exit Function

	s=Split(sFilter_Plugin_TUser_LoadInfobyID,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Level,Password,Email,HomePage,Count,Alias,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TUser_LoadInfoByArray
'参数:ID,Name,Level,Password,Email,HomePage,Count,Alias
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TUser_LoadInfoByArray
Function Filter_Plugin_TUser_LoadInfoByArray(ByRef ID,ByRef Name,ByRef Level,ByRef Password,ByRef Email,ByRef HomePage,ByRef Count,ByRef Alias,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TUser_LoadInfoByArray="" Then Exit Function

	s=Split(sFilter_Plugin_TUser_LoadInfoByArray,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Level,Password,Email,HomePage,Count,Alias,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TUser_Edit
'参数:ID,Name,Level,Password,Email,HomePage,Count,Alias
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TUser_Edit
Function Filter_Plugin_TUser_Edit(ByRef ID,ByRef Name,ByRef Level,ByRef Password,ByRef Email,ByRef HomePage,ByRef Count,ByRef Alias,ByRef MetaString,ByRef currentUser)

	Dim s,i

	If sFilter_Plugin_TUser_Edit="" Then Exit Function

	s=Split(sFilter_Plugin_TUser_Edit,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Level,Password,Email,HomePage,Count,Alias,MetaString,currentUser")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TUser_Register
'参数:ID,Name,Level,Password,Email,HomePage,Count,Alias
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TUser_Register
Function Filter_Plugin_TUser_Register(ByRef ID,ByRef Name,ByRef Level,ByRef Password,ByRef Email,ByRef HomePage,ByRef Count,ByRef Alias,ByRef MetaString,ByRef currentUser)

	Dim s,i

	If sFilter_Plugin_TUser_Register="" Then Exit Function

	s=Split(sFilter_Plugin_TUser_Register,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Level,Password,Email,HomePage,Count,Alias,MetaString,currentUser")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TUser_Del
'参数:ID,Name,Level,Password,Email,HomePage,Count,Alias
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TUser_Del
Function Filter_Plugin_TUser_Del(ByRef ID,ByRef Name,ByRef Level,ByRef Password,ByRef Email,ByRef HomePage,ByRef Count,ByRef Alias,ByRef MetaString,ByRef currentUser)

	Dim s,i

	If sFilter_Plugin_TUser_Del="" Then Exit Function

	s=Split(sFilter_Plugin_TUser_Del,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Level,Password,Email,HomePage,Count,Alias,MetaString,currentUser")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TComment_Post
'参数:ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,ParentID,ParentCount
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TComment_Post
Function Filter_Plugin_TComment_Post(ByRef ID,ByRef log_ID,ByRef AuthorID,ByRef Author,ByRef Content,ByRef Email,ByRef HomePage,ByRef PostTime,ByRef IP,ByRef Agent,ByRef Reply,ByRef LastReplyIP,ByRef LastReplyTime,ByRef ParentID,ByRef ParentCount,ByRef IsCheck,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TComment_Post="" Then Exit Function

	s=Split(sFilter_Plugin_TComment_Post,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,ParentCount,IsCheck,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TComment_LoadInfoByArray
'参数:ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,ParentID,ParentCount
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TComment_LoadInfoByArray
Function Filter_Plugin_TComment_LoadInfoByArray(ByRef ID,ByRef log_ID,ByRef AuthorID,ByRef Author,ByRef Content,ByRef Email,ByRef HomePage,ByRef PostTime,ByRef IP,ByRef Agent,ByRef Reply,ByRef LastReplyIP,ByRef LastReplyTime,ByRef ParentID,ByRef ParentCount,ByRef IsCheck,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TComment_LoadInfoByArray="" Then Exit Function

	s=Split(sFilter_Plugin_TComment_LoadInfoByArray,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,ParentCount,IsCheck,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TComment_Del
'参数:ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,ParentID,ParentCount
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TComment_Del
Function Filter_Plugin_TComment_Del(ByRef ID,ByRef log_ID,ByRef AuthorID,ByRef Author,ByRef Content,ByRef Email,ByRef HomePage,ByRef PostTime,ByRef IP,ByRef Agent,ByRef Reply,ByRef LastReplyIP,ByRef LastReplyTime,ByRef ParentID,ByRef ParentCount,ByRef IsCheck,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TComment_Del="" Then Exit Function

	s=Split(sFilter_Plugin_TComment_Del,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,ParentCount,IsCheck,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TComment_LoadInfoByID
'参数:ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,ParentID,ParentCount
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TComment_LoadInfoByID
Function Filter_Plugin_TComment_LoadInfoByID(ByRef ID,ByRef log_ID,ByRef AuthorID,ByRef Author,ByRef Content,ByRef Email,ByRef HomePage,ByRef PostTime,ByRef IP,ByRef Agent,ByRef Reply,ByRef LastReplyIP,ByRef LastReplyTime,ByRef ParentID,ByRef ParentCount,ByRef IsCheck,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TComment_LoadInfoByID="" Then Exit Function

	s=Split(sFilter_Plugin_TComment_LoadInfoByID,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,log_ID,AuthorID,Author,Content,Email,HomePage,PostTime,IP,Agent,Reply,LastReplyIP,LastReplyTime,ParentID,ParentCount,IsCheck,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TComment_MakeTemplate_Template
'参数:html
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TComment_MakeTemplate_Template
Function Filter_Plugin_TComment_MakeTemplate_Template(ByRef html)

	Dim s,i

	If sFilter_Plugin_TComment_MakeTemplate_Template="" Then Exit Function

	s=Split(sFilter_Plugin_TComment_MakeTemplate_Template,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "html")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TComment_MakeTemplate_TemplateTags
'参数:aryTemplateTagsName,aryTemplateTagsValue
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TComment_MakeTemplate_TemplateTags
Function Filter_Plugin_TComment_MakeTemplate_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)

	Dim s,i

	If sFilter_Plugin_TComment_MakeTemplate_TemplateTags="" Then Exit Function

	s=Split(sFilter_Plugin_TComment_MakeTemplate_TemplateTags,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "aryTemplateTagsName,aryTemplateTagsValue")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTrackBack_Post
'参数:ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTrackBack_Post
Function Filter_Plugin_TTrackBack_Post(ByRef ID,ByRef log_ID,ByRef URL,ByRef Title,ByRef Blog,ByRef Excerpt,ByRef PostTime,ByRef IP,ByRef Agent,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TTrackBack_Post="" Then Exit Function

	s=Split(sFilter_Plugin_TTrackBack_Post,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTrackBack_LoadInfoByArray
'参数:ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTrackBack_LoadInfoByArray
Function Filter_Plugin_TTrackBack_LoadInfoByArray(ByRef ID,ByRef log_ID,ByRef URL,ByRef Title,ByRef Blog,ByRef Excerpt,ByRef PostTime,ByRef IP,ByRef Agent,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TTrackBack_LoadInfoByArray="" Then Exit Function

	s=Split(sFilter_Plugin_TTrackBack_LoadInfoByArray,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTrackBack_Del
'参数:ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTrackBack_Del
Function Filter_Plugin_TTrackBack_Del(ByRef ID,ByRef log_ID,ByRef URL,ByRef Title,ByRef Blog,ByRef Excerpt,ByRef PostTime,ByRef IP,ByRef Agent,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TTrackBack_Del="" Then Exit Function

	s=Split(sFilter_Plugin_TTrackBack_Del,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTrackBack_LoadInfoByID
'参数:ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTrackBack_LoadInfoByID
Function Filter_Plugin_TTrackBack_LoadInfoByID(ByRef ID,ByRef log_ID,ByRef URL,ByRef Title,ByRef Blog,ByRef Excerpt,ByRef PostTime,ByRef IP,ByRef Agent,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TTrackBack_LoadInfoByID="" Then Exit Function

	s=Split(sFilter_Plugin_TTrackBack_LoadInfoByID,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,log_ID,URL,Title,Blog,Excerpt,PostTime,IP,Agent,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTrackBack_MakeTemplate_Template
'参数:html
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTrackBack_MakeTemplate_Template
Function Filter_Plugin_TTrackBack_MakeTemplate_Template(ByRef html)

	Dim s,i

	If sFilter_Plugin_TTrackBack_MakeTemplate_Template="" Then Exit Function

	s=Split(sFilter_Plugin_TTrackBack_MakeTemplate_Template,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "html")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTrackBack_MakeTemplate_TemplateTags
'参数:aryTemplateTagsName,aryTemplateTagsValue
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTrackBack_MakeTemplate_TemplateTags
Function Filter_Plugin_TTrackBack_MakeTemplate_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)

	Dim s,i

	If sFilter_Plugin_TTrackBack_MakeTemplate_TemplateTags="" Then Exit Function

	s=Split(sFilter_Plugin_TTrackBack_MakeTemplate_TemplateTags,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "aryTemplateTagsName,aryTemplateTagsValue")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTag_Post
'参数:ID,Name,Intro,Order,Count
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTag_Post
Function Filter_Plugin_TTag_Post(ByRef ID,ByRef Name,ByRef Intro,ByRef Order,ByRef Count,ByRef ParentID,ByRef Alias,ByRef TemplateName,ByRef FullUrl,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TTag_Post="" Then Exit Function

	s=Split(sFilter_Plugin_TTag_Post,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTag_LoadInfoByArray
'参数:ID,Name,Intro,Order,Count
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTag_LoadInfoByArray
Function Filter_Plugin_TTag_LoadInfoByArray(ByRef ID,ByRef Name,ByRef Intro,ByRef Order,ByRef Count,ByRef ParentID,ByRef Alias,ByRef TemplateName,ByRef FullUrl,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TTag_LoadInfoByArray="" Then Exit Function

	s=Split(sFilter_Plugin_TTag_LoadInfoByArray,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTag_Del
'参数:ID,Name,Intro,Order,Count
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTag_Del
Function Filter_Plugin_TTag_Del(ByRef ID,ByRef Name,ByRef Intro,ByRef Order,ByRef Count,ByRef ParentID,ByRef Alias,ByRef TemplateName,ByRef FullUrl,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TTag_Del="" Then Exit Function

	s=Split(sFilter_Plugin_TTag_Del,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTag_LoadInfoByID
'参数:ID,Name,Intro,Order,Count
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTag_LoadInfoByID
Function Filter_Plugin_TTag_LoadInfoByID(ByRef ID,ByRef Name,ByRef Intro,ByRef Order,ByRef Count,ByRef ParentID,ByRef Alias,ByRef TemplateName,ByRef FullUrl,ByRef MetaString)

	Dim s,i

	If sFilter_Plugin_TTag_LoadInfoByID="" Then Exit Function

	s=Split(sFilter_Plugin_TTag_LoadInfoByID,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,Name,Intro,Order,Count,ParentID,Alias,TemplateName,FullUrl,MetaString")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTag_MakeTemplate_Template
'参数:html
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTag_MakeTemplate_Template
Function Filter_Plugin_TTag_MakeTemplate_Template(ByRef html)

	Dim s,i

	If sFilter_Plugin_TTag_MakeTemplate_Template="" Then Exit Function

	s=Split(sFilter_Plugin_TTag_MakeTemplate_Template,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "html")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTag_MakeTemplate_TemplateTags
'参数:aryTemplateTagsName,aryTemplateTagsValue
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTag_MakeTemplate_TemplateTags
Function Filter_Plugin_TTag_MakeTemplate_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)

	Dim s,i

	If sFilter_Plugin_TTag_MakeTemplate_TemplateTags="" Then Exit Function

	s=Split(sFilter_Plugin_TTag_MakeTemplate_TemplateTags,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "aryTemplateTagsName,aryTemplateTagsValue")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TUpLoadFile_UpLoad
'参数:ID,AuthorID,FileSize,FileName,PostTime,DirByTime
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TUpLoadFile_UpLoad
Function Filter_Plugin_TUpLoadFile_UpLoad(ByRef ID,ByRef AuthorID,ByRef FileSize,ByRef FileName,ByRef PostTime,ByRef FileIntro,ByRef DirByTime,ByRef Quote,ByRef Meta)

	Dim s,i

	If sFilter_Plugin_TUpLoadFile_UpLoad="" Then Exit Function

	s=Split(sFilter_Plugin_TUpLoadFile_UpLoad,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Quote,Meta")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TUpLoadFile_LoadInfoByArray
'参数:ID,AuthorID,FileSize,FileName,PostTime,DirByTime
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TUpLoadFile_LoadInfoByArray
Function Filter_Plugin_TUpLoadFile_LoadInfoByArray(ByRef ID,ByRef AuthorID,ByRef FileSize,ByRef FileName,ByRef PostTime,ByRef FileIntro,ByRef DirByTime,ByRef Quote,ByRef Meta)

	Dim s,i

	If sFilter_Plugin_TUpLoadFile_LoadInfoByArray="" Then Exit Function

	s=Split(sFilter_Plugin_TUpLoadFile_LoadInfoByArray,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Quote,Meta")
	Next

End Function




'**************************************************<
'类型:filter
'名称:Filter_Plugin_TUpLoadFile_Del
'参数:ID,AuthorID,FileSize,FileName,PostTime,DirByTime
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TUpLoadFile_Del
Function Filter_Plugin_TUpLoadFile_Del(ByRef ID,ByRef AuthorID,ByRef FileSize,ByRef FileName,ByRef PostTime,ByRef FileIntro,ByRef DirByTime,ByRef Quote,ByRef Meta)

	Dim s,i

	If sFilter_Plugin_TUpLoadFile_Del="" Then Exit Function

	s=Split(sFilter_Plugin_TUpLoadFile_Del,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Quote,Meta")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TUpLoadFile_LoadInfoByID
'参数:ID,AuthorID,FileSize,FileName,PostTime,DirByTime
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TUpLoadFile_LoadInfoByID
Function Filter_Plugin_TUpLoadFile_LoadInfoByID(ByRef ID,ByRef AuthorID,ByRef FileSize,ByRef FileName,ByRef PostTime,ByRef FileIntro,ByRef DirByTime,ByRef Quote,ByRef Meta)

	Dim s,i

	If sFilter_Plugin_TUpLoadFile_LoadInfoByID="" Then Exit Function

	s=Split(sFilter_Plugin_TUpLoadFile_LoadInfoByID,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "ID,AuthorID,FileSize,FileName,PostTime,FileIntro,DirByTime,Quote,Meta")
	Next

End Function




'**************************************************<
'类型:action
'名称:Action_Plugin_System_Initialize
'参数:无
'说明:在系统初始化时被调用
'**************************************************>
Dim Action_Plugin_TTag_Url()
ReDim Action_Plugin_TTag_Url(0)
Dim bAction_Plugin_TTag_Url
Dim sAction_Plugin_TTag_Url




'**************************************************<
'类型:filter
'名称:Filter_Plugin_TTag_Url
'参数:Url
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TTag_Url
Function Filter_Plugin_TTag_Url(ByRef Url)

	Dim s,i

	If sFilter_Plugin_TTag_Url="" Then Exit Function

	s=Split(sFilter_Plugin_TTag_Url,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "url")
	Next

End Function




'**************************************************<
'类型:action
'名称:Action_Plugin_System_Initialize
'参数:无
'说明:在系统初始化时被调用
'**************************************************>
Dim Action_Plugin_TCategory_Url()
ReDim Action_Plugin_TCategory_Url(0)
Dim bAction_Plugin_TCategory_Url
Dim sAction_Plugin_TCategory_Url





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TCategory_Url
'参数:Url
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TCategory_Url
Function Filter_Plugin_TCategory_Url(ByRef Url)

	Dim s,i

	If sFilter_Plugin_TCategory_Url="" Then Exit Function

	s=Split(sFilter_Plugin_TCategory_Url,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "url")
	Next

End Function




'**************************************************<
'类型:action
'名称:Action_Plugin_System_Initialize
'参数:无
'说明:在系统初始化时被调用
'**************************************************>
Dim Action_Plugin_TArticle_Url()
ReDim Action_Plugin_TArticle_Url(0)
Dim bAction_Plugin_TArticle_Url
Dim sAction_Plugin_TArticle_Url




'**************************************************<
'类型:filter
'名称:Filter_Plugin_TArticle_Url
'参数:Url
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TArticle_Url
Function Filter_Plugin_TArticle_Url(ByRef Url)

	Dim s,i

	If sFilter_Plugin_TArticle_Url="" Then Exit Function

	s=Split(sFilter_Plugin_TArticle_Url,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "url")
	Next

End Function




'**************************************************<
'类型:action
'名称:Action_Plugin_System_Initialize
'参数:无
'说明:在系统初始化时被调用
'**************************************************>
Dim Action_Plugin_TUser_Url()
ReDim Action_Plugin_TUser_Url(0)
Dim bAction_Plugin_TUser_Url
Dim sAction_Plugin_TUser_Url




'**************************************************<
'类型:filter
'名称:Filter_Plugin_TUser_Url
'参数:Url
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TUser_Url
Function Filter_Plugin_TUser_Url(ByRef Url)

	Dim s,i

	If sFilter_Plugin_TUser_Url="" Then Exit Function

	s=Split(sFilter_Plugin_TUser_Url,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "url")
	Next

End Function







'***************
'3.response
'***************


'**************************************************<
'类型:response
'名称:Response_Plugin_ArticleMng_SubMenu
'参数:无
'说明:文章管理子菜单
'**************************************************>
Dim Response_Plugin_ArticleMng_SubMenu
Response_Plugin_ArticleMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_CategoryMng_SubMenu
'参数:无
'说明:分类管理子菜单
'**************************************************>
Dim Response_Plugin_CategoryMng_SubMenu
Response_Plugin_CategoryMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_CommentMng_SubMenu
'参数:无
'说明:评论管理子菜单
'**************************************************>
Dim Response_Plugin_CommentMng_SubMenu
Response_Plugin_CommentMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_TrackBackMng_SubMenu
'参数:无
'说明:引用管理子菜单
'**************************************************>
Dim Response_Plugin_TrackBackMng_SubMenu
Response_Plugin_TrackBackMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_UserMng_SubMenu
'参数:无
'说明:用户管理子菜单
'**************************************************>
Dim Response_Plugin_UserMng_SubMenu
Response_Plugin_UserMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_FileMng_SubMenu
'参数:无
'说明:附件管理子菜单
'**************************************************>
Dim Response_Plugin_FileMng_SubMenu
Response_Plugin_FileMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_TagMng_SubMenu
'参数:无
'说明:Tags管理子菜单
'**************************************************>
Dim Response_Plugin_TagMng_SubMenu
Response_Plugin_TagMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_PlugInMng_SubMenu
'参数:无
'说明:插件管理子菜单
'**************************************************>
Dim Response_Plugin_PlugInMng_SubMenu
Response_Plugin_PlugInMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_SiteInfo_SubMenu
'参数:无
'说明:后台首页管理子菜单
'**************************************************>
Dim Response_Plugin_SiteInfo_SubMenu
Response_Plugin_SiteInfo_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_SiteFileMng_SubMenu
'参数:无
'说明:站内文件管理子菜单
'**************************************************>
Dim Response_Plugin_SiteFileMng_SubMenu
Response_Plugin_SiteFileMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_SiteFileEdt_SubMenu
'参数:无
'说明:站内文件编辑子菜单
'**************************************************>
Dim Response_Plugin_SiteFileEdt_SubMenu
Response_Plugin_SiteFileEdt_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_AskFileReBuild_SubMenu
'参数:无
'说明:请求文章重建子菜单
'**************************************************>
Dim Response_Plugin_AskFileReBuild_SubMenu
Response_Plugin_AskFileReBuild_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_ThemeMng_SubMenu
'参数:无
'说明:主题管理子菜单
'**************************************************>
Dim Response_Plugin_ThemeMng_SubMenu
Response_Plugin_ThemeMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_SettingMng_SubMenu
'参数:无
'说明:网站设置子菜单
'**************************************************>
Dim Response_Plugin_SettingMng_SubMenu
Response_Plugin_SettingMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_LinkMng_SubMenu
'参数:无
'说明:链接管理子菜单
'**************************************************>
Dim Response_Plugin_LinkMng_SubMenu
Response_Plugin_LinkMng_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_ArticleEdt_SubMenu
'参数:无
'说明:文件编辑页子菜单
'**************************************************>
Dim Response_Plugin_ArticleEdt_SubMenu
Response_Plugin_ArticleEdt_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_CategoryEdt_SubMenu
'参数:无
'说明:分类编辑子菜单
'**************************************************>
Dim Response_Plugin_CategoryEdt_SubMenu
Response_Plugin_CategoryEdt_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_CommentEdt_SubMenu
'参数:无
'说明:评论管理子菜单
'**************************************************>
Dim Response_Plugin_CommentEdt_SubMenu
Response_Plugin_CommentEdt_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_TagEdt_SubMenu
'参数:无
'说明:Tags编辑子菜单
'**************************************************>
Dim Response_Plugin_TagEdt_SubMenu
Response_Plugin_TagEdt_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_UserEdt_SubMenu
'参数:无
'说明:用户编辑子菜单
'**************************************************>
Dim Response_Plugin_UserEdt_SubMenu
Response_Plugin_UserEdt_SubMenu=""


'**************************************************<
'类型:response
'名称:Response_Plugin_Edit_HtmlTag
'参数:无
'说明:文件编辑页HTML标签
'**************************************************>
Dim Response_Plugin_Edit_HtmlTag
Response_Plugin_Edit_HtmlTag=""


'**************************************************<
'类型:response
'名称:Response_Plugin_Edit_UbbTag
'参数:无
'说明:文件编辑页UBB标签
'**************************************************>
Dim Response_Plugin_Edit_UbbTag
Response_Plugin_Edit_UbbTag=""


'**************************************************<
'类型:response
'名称:Response_Plugin_Edit_Form
'参数:无
'说明:文件编辑页Form标签
'**************************************************>
Dim Response_Plugin_Edit_Form
Response_Plugin_Edit_Form=""



'**************************************************<
'类型:response
'名称:Response_Plugin_Edit_Form2
'参数:无
'说明:文件编辑页Form2标签
'**************************************************>
Dim Response_Plugin_Edit_Form2
Response_Plugin_Edit_Form2=""






'**************************************************<
'类型:action
'名称:Action_Plugin_BuildAllCache_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BuildAllCache_Begin()
ReDim Action_Plugin_BuildAllCache_Begin(0)
Dim bAction_Plugin_BuildAllCache_Begin
Dim sAction_Plugin_BuildAllCache_Begin




'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Calendar_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_Calendar_Begin()
ReDim Action_Plugin_BlogReBuild_Calendar_Begin(0)
Dim bAction_Plugin_BlogReBuild_Calendar_Begin
Dim sAction_Plugin_BlogReBuild_Calendar_Begin




'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Archives_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_Archives_Begin()
ReDim Action_Plugin_BlogReBuild_Archives_Begin(0)
Dim bAction_Plugin_BlogReBuild_Archives_Begin
Dim sAction_Plugin_BlogReBuild_Archives_Begin




'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Catalogs_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_Catalogs_Begin()
ReDim Action_Plugin_BlogReBuild_Catalogs_Begin(0)
Dim bAction_Plugin_BlogReBuild_Catalogs_Begin
Dim sAction_Plugin_BlogReBuild_Catalogs_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Categorys_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_Categorys_Begin()
ReDim Action_Plugin_BlogReBuild_Categorys_Begin(0)
Dim bAction_Plugin_BlogReBuild_Categorys_Begin
Dim sAction_Plugin_BlogReBuild_Categorys_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Authors_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_Authors_Begin()
ReDim Action_Plugin_BlogReBuild_Authors_Begin(0)
Dim bAction_Plugin_BlogReBuild_Authors_Begin
Dim sAction_Plugin_BlogReBuild_Authors_Begin




'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Tags_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_Tags_Begin()
ReDim Action_Plugin_BlogReBuild_Tags_Begin(0)
Dim bAction_Plugin_BlogReBuild_Tags_Begin
Dim sAction_Plugin_BlogReBuild_Tags_Begin




'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Previous_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_Previous_Begin()
ReDim Action_Plugin_BlogReBuild_Previous_Begin(0)
Dim bAction_Plugin_BlogReBuild_Previous_Begin
Dim sAction_Plugin_BlogReBuild_Previous_Begin





'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Comments_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_Comments_Begin()
ReDim Action_Plugin_BlogReBuild_Comments_Begin(0)
Dim bAction_Plugin_BlogReBuild_Comments_Begin
Dim sAction_Plugin_BlogReBuild_Comments_Begin




'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_GuestComments_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_GuestComments_Begin()
ReDim Action_Plugin_BlogReBuild_GuestComments_Begin(0)
Dim bAction_Plugin_BlogReBuild_GuestComments_Begin
Dim sAction_Plugin_BlogReBuild_GuestComments_Begin





'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_TrackBacks_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_TrackBacks_Begin()
ReDim Action_Plugin_BlogReBuild_TrackBacks_Begin(0)
Dim bAction_Plugin_BlogReBuild_TrackBacks_Begin
Dim sAction_Plugin_BlogReBuild_TrackBacks_Begin




'**************************************************<
'类型:action
'名称:Action_Plugin_BlogReBuild_Statistics_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_BlogReBuild_Statistics_Begin()
ReDim Action_Plugin_BlogReBuild_Statistics_Begin(0)
Dim bAction_Plugin_BlogReBuild_Statistics_Begin
Dim sAction_Plugin_BlogReBuild_Statistics_Begin




'**************************************************<
'类型:action
'名称:Action_Plugin_ExportRSS_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_ExportRSS_Begin()
ReDim Action_Plugin_ExportRSS_Begin(0)
Dim bAction_Plugin_ExportRSS_Begin
Dim sAction_Plugin_ExportRSS_Begin




'**************************************************<
'类型:action
'名称:Action_Plugin_ExportATOM_Begin
'参数:无
'说明:c_system_base.asp
'**************************************************>
Dim Action_Plugin_ExportATOM_Begin()
ReDim Action_Plugin_ExportATOM_Begin(0)
Dim bAction_Plugin_ExportATOM_Begin
Dim sAction_Plugin_ExportATOM_Begin





'**************************************************<
'类型:action
'名称:Action_Plugin_TGuestBook_Export_Begin
'参数:无
'说明:TGuestBook_Export_Begin
'**************************************************>
Dim Action_Plugin_TGuestBook_Export_Begin()
ReDim Action_Plugin_TGuestBook_Export_Begin(0)
Dim bAction_Plugin_TGuestBook_Export_Begin
Dim sAction_Plugin_TGuestBook_Export_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_TGuestBook_Export_End
'参数:无
'说明:TGuestBook_Export_End
'**************************************************>
Dim Action_Plugin_TGuestBook_Export_End()
ReDim Action_Plugin_TGuestBook_Export_End(0)
Dim bAction_Plugin_TGuestBook_Export_End
Dim sAction_Plugin_TGuestBook_Export_End





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TGuestBook_Export_TemplateTags
'参数:aryTemplateTagsName,aryTemplateTagsValue
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TGuestBook_Export_TemplateTags
Function Filter_Plugin_TGuestBook_Export_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)

	Dim s,i

	If sFilter_Plugin_TGuestBook_Export_TemplateTags="" Then Exit Function

	s=Split(sFilter_Plugin_TGuestBook_Export_TemplateTags,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "aryTemplateTagsName,aryTemplateTagsValue")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TGuestBook_Build_TemplateTags
'参数:aryTemplateTagsName,aryTemplateTagsValue
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TGuestBook_Build_TemplateTags
Function Filter_Plugin_TGuestBook_Build_TemplateTags(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)

	Dim s,i

	If sFilter_Plugin_TGuestBook_Build_TemplateTags="" Then Exit Function

	s=Split(sFilter_Plugin_TGuestBook_Build_TemplateTags,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "aryTemplateTagsName,aryTemplateTagsValue")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TGuestBook_Build_Template
'参数:html,wapHtml
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TGuestBook_Build_Template
Function Filter_Plugin_TGuestBook_Build_Template(ByRef html,ByRef wapHtml)

	Dim s,i

	If sFilter_Plugin_TGuestBook_Build_Template="" Then Exit Function

	s=Split(sFilter_Plugin_TGuestBook_Build_Template,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "html,wapHtml")
	Next

End Function





'**************************************************<
'类型:filter
'名称:Filter_Plugin_TGuestBook_Build_Template_Sub
'参数:Template_Article_Single
'说明:
'调用:
'**************************************************>
Dim sFilter_Plugin_TGuestBook_Build_Template_Sub
Function Filter_Plugin_TGuestBook_Build_Template_Sub(ByRef Template_Article_Single)

	Dim s,i

	If sFilter_Plugin_TGuestBook_Build_Template_Sub="" Then Exit Function

	s=Split(sFilter_Plugin_TGuestBook_Build_Template_Sub,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "Template_Article_Single")
	Next

End Function




'**************************************************<
'类型:action
'名称:Action_Plugin_Tags_Begin
'参数:无
'说明:tags.asp
'**************************************************>
Dim Action_Plugin_Tags_Begin()
ReDim Action_Plugin_Tags_Begin(0)
Dim bAction_Plugin_Tags_Begin
Dim sAction_Plugin_Tags_Begin



'**************************************************<
'类型:action
'名称:Action_Plugin_Tags_End
'参数:无
'说明:tags.asp
'**************************************************>
Dim Action_Plugin_Tags_End()
ReDim Action_Plugin_Tags_End(0)
Dim bAction_Plugin_Tags_End
Dim sAction_Plugin_Tags_End


'**************************************************<
'类型:response
'名称:Response_Plugin_AdminLeft_Plugin
'参数:无
'说明:左侧菜单
'**************************************************>
Dim Response_Plugin_AdminLeft_Plugin
Response_Plugin_AdminLeft_Plugin=""


'以下为uEditor的所有接口：
Dim Action_Plugin_UEditor_FileUpload_Begin()
ReDim Action_Plugin_UEditor_FileUpload_Begin(0)
Dim bAction_Plugin_UEditor_FileUpload_Begin
Dim sAction_Plugin_UEditor_FileUpload_Begin

Dim Action_Plugin_UEditor_FileUpload_End()
ReDim Action_Plugin_UEditor_FileUpload_End(0)
Dim bAction_Plugin_UEditor_FileUpload_End
Dim sAction_Plugin_UEditor_FileUpload_End

Dim Action_Plugin_UEditor_imageManager_Begin()
ReDim Action_Plugin_UEditor_imageManager_Begin(0)
Dim bAction_Plugin_UEditor_imageManager_Begin
Dim sAction_Plugin_UEditor_imageManager_Begin

Dim Action_Plugin_UEditor_imageManager_End()
ReDim Action_Plugin_UEditor_imageManager_End(0)
Dim bAction_Plugin_UEditor_imageManager_End
Dim sAction_Plugin_UEditor_imageManager_End

Dim Action_Plugin_UEditor_Config_Begin()
ReDim Action_Plugin_UEditor_Config_Begin(0)
Dim bAction_Plugin_UEditor_Config_Begin
Dim sAction_Plugin_UEditor_Config_Begin

Dim Action_Plugin_UEditor_Config_End()
ReDim Action_Plugin_UEditor_Config_End(0)
Dim bAction_Plugin_UEditor_Config_End
Dim sAction_Plugin_UEditor_Config_End

Dim Action_Plugin_UEditor_getRemoteImage_Begin()
ReDim Action_Plugin_UEditor_getRemoteImage_Begin(0)
Dim bAction_Plugin_UEditor_getRemoteImage_Begin
Dim sAction_Plugin_UEditor_getRemoteImage_Begin

Dim Action_Plugin_UEditor_getRemoteImage_End()
ReDim Action_Plugin_UEditor_getRemoteImage_End(0)
Dim bAction_Plugin_UEditor_getRemoteImage_End
Dim sAction_Plugin_UEditor_getRemoteImage_End

Dim Action_Plugin_UEditor_getmovie_Begin()
ReDim Action_Plugin_UEditor_getmovie_Begin(0)
Dim bAction_Plugin_UEditor_getmovie_Begin
Dim sAction_Plugin_UEditor_getmovie_Begin

Dim Action_Plugin_UEditor_getmovie_End()
ReDim Action_Plugin_UEditor_getmovie_End(0)
Dim bAction_Plugin_UEditor_getmovie_End
Dim sAction_Plugin_UEditor_getmovie_End

Dim Action_Plugin_UEditor_getcontent_Begin()
ReDim Action_Plugin_UEditor_getcontent_Begin(0)
Dim bAction_Plugin_UEditor_getcontent_Begin
Dim sAction_Plugin_UEditor_getcontent_Begin

Dim Action_Plugin_UEditor_getcontent_End()
ReDim Action_Plugin_UEditor_getcontent_End(0)
Dim bAction_Plugin_UEditor_getcontent_End
Dim sAction_Plugin_UEditor_getcontent_End



Dim sFilter_Plugin_UEditor_Config
Function Filter_Plugin_UEditor_Config(ByRef strJSContent)

	Dim s,i

	If sFilter_Plugin_UEditor_Config="" Then Exit Function

	s=Split(sFilter_Plugin_UEditor_Config,"|")

	For i=0 To UBound(s)-1
		Call Execute(s(i) & " " & "strJSContent")
	Next

End Function


%>
