<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Devo Or Newer
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    批量管理文章插件 - 跳转页
'// 最后修改：   2008-10-24
'// 最后版本:    1.4
'///////////////////////////////////////////////////////////////////////////////

'注册插件
Call RegisterPlugin("BatchArticles","ActivePlugin_BatchArticles")

Function ActivePlugin_BatchArticles() 

	'挂上接口
	Call Add_Action_Plugin("Action_Plugin_ArticleMng_Begin","If Request.QueryString(""type"")<>""Page"" Then Response.Redirect ""../zb_users/plugin/BatchArticles/main.asp""")

End Function

'安装插件
Function InstallPlugin_BatchArticles()
	BatchArticles_Init
	BatchArticles_ShowWarning()
	Call SetBlogHint_Custom("? 提示:批量管理文章已启用, 此插件将替换博客后台的文章管理.")

End Function

'卸载插件
Function UnInstallPlugin_BatchArticles()

	BatchArticles_ShowWarning()
	Call SetBlogHint_Custom("? 提示:批量管理文章已停用, 博客后台的文章管理可恢复使用.")

End Function

Function BatchArticles_SkipWarning()
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load "BatchArticles"
	objConfig.Write "ShowWarning",False
	objConfig.Save
End Function

Function BatchArticles_ShowWarning()
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load "BatchArticles"
	objConfig.Write "ShowWarning",True
	objConfig.Save
End Function

Sub BatchArticles_Init()
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load "BatchArticles"
	If objConfig.Exists("Version")=False Then
		objConfig.Write "ShowWarning","True"
		objConfig.Write "UseTagMng","False"
		objConfig.Write "UseTagCloud","False"
		objConfig.Write "UseTagHint","False"
		objConfig.Write "Version","1.0"
		objConfig.Save
	End If
End Sub
%>