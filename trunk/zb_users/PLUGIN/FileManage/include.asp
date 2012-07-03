<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.9 其它版本的Z-blog未知
'// 插件制作:    ZSXSOFT(http://www.zsxsoft.com/)
'// 备    注:    FileManage - 挂口函数页
'///////////////////////////////////////////////////////////////////////////////

'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************

'注册插件
Call RegisterPlugin("FileManage","ActivePlugin_FileManage")

'挂口部分
Function ActivePlugin_FileManage()

	Call Add_Action_Plugin("Action_Plugin_Admin_Begin","FileManage_Include")
	Call Add_AdminLeft_Plugin("文件管理","http://www.zsxsoft.com")

End Function
'*********************************************************
'直接接管文件管理
'*********************************************************
Sub FileManage_Include
	Dim strAct
	strAct=Request.QueryString("act")
	If Not CheckRights(strAct) Then Call ShowError(6)
	Select Case Request.QueryString("act")

		Case "SiteFileMng" Response.Redirect ZC_BLOG_HOST & "/zb_users/PLUGIN/FileManage/main.asp?act=SiteFileMng&path="&Server.URLEncode(Request.QueryString("path"))&"&opath="&Server.URLEncode(Request.QueryString("opath")):Response.End
		Case "SiteFileEdt" Response.Redirect ZC_BLOG_HOST & "/zb_users/PLUGIN/FileManage/main.asp?act=SiteFileEdt&path="&Server.URLEncode(Request.QueryString("path"))&"&opath="&Server.URLEncode(Request.QueryString("opath")):Response.End

	End Select
End Sub

%>