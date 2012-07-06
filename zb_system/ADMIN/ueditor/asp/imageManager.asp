<!--#include file="up_inc.asp"-->
<!-- #include file="../../../../zb_users\c_option.asp" -->
<!-- #include file="../../../function\c_function.asp" -->
<!-- #include file="../../../function\c_function_md5.asp" -->
<!-- #include file="../../../function\c_system_lib.asp" -->
<!-- #include file="../../../function\c_system_base.asp" -->
<!-- #include file="../../../function\c_system_event.asp" -->
<!-- #include file="../../../function\c_system_plugin.asp" -->
<!-- #include file="../../../function\rss_lib.asp" -->
<!-- #include file="../../../function\atom_lib.asp" -->
<!-- #include file="../../../../zb_users\plugin\p_config.asp" -->
<%
On Error Resume Next
Call System_Initialize()
Call CheckReference("")
If Not CheckRights("ArticleEdt") Then Call ShowError(6)

For Each sAction_Plugin_imageManager_Begin in Action_Plugin_imageManager_Begin
	If Not IsEmpty(sAction_Plugin_imageManager_Begin) Then Call Execute(sAction_Plugin_imageManager_Begin)
Next
	Dim strResponse
	strResponse="noimage.jpg"
	'strResponse="此功能(imagemanager.asp)系统默认不开放，请安装必要插件。"

For Each sAction_Plugin_imageManager_End in Action_Plugin_imageManager_End
	If Not IsEmpty(sAction_Plugin_imageManager_End) Then Call Execute(sAction_Plugin_imageManager_End)
Next
	Response.Write strResponse

%>