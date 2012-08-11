<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    myllop-大猪
'// 版权所有:    www.izhu.org
'// 技术支持:    myllop#gmail.com
'// 程序名称:    大猪滔滔插件
'// 英文名称:    dztaotao
'// 开始时间:    2012-8-10
'// 最后修改:    
'// 备    注:    only for zblog2.0
'///////////////////////////////////////////////////////////////////////////////

Dim DZTAOTAO_TITLE_VALUE
Dim DZTAOTAO_RELEASE_VALUE
Dim DZTAOTAO_PAGECOUNT_VALUE
Dim DZTAOTAO_PAGEWIDTH_VALUE
Dim DZTAOTAO_CHK_VALUE
Dim DZTAOTAO_CMTCHK_VALUE
Dim DZTAOTAO_CMTLIMIT_VALUE
Dim DZTAOTAO_ISIMG_VALUE

Dim dztaotao_Config
	
	
	
	

'注册插件
Call RegisterPlugin("dztaotao","ActivePlugin_dztaotao")

Function ActivePlugin_dztaotao() 

	'挂上接口
	'Call Add_Action_Plugin("Action_Plugin_MakeBlogReBuild_Begin","Call dztaotao_CreateDB")
	
	Call Add_Action_Plugin("Action_Plugin_MakeBlogReBuild_Begin","Call dztaotao_CreateDB")
	
	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(5,"大猪滔滔",GetCurrentHost&"zb_users/plugin/dztaotao/setting.asp","nav_dztaotao","adztaotao",GetCurrentHost&"zb_users/plugin/dztaotao/images/navlogo_1.png"))
	
	'网站管理加上二级菜单项
    Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu", MakeSubMenu("淘淘设置", GetCurrentHost&"zb_users/plugin/dztaotao/setting.asp", "m-left", FALSE))
    Call Add_Response_Plugin("Response_Plugin_SiteInfo_SubMenu",MakeSubMenu("[淘淘管理]",GetCurrentHost&"/zb_users/plugin/dztaotao/admin.asp?a=list&page=1","m-left",False))


End Function


'更新配置数据
function dztaotao_UpdateConfig()

	Set dztaotao_Config = New TConfig
	dztaotao_Config.Load("dztaotao")
	If dztaotao_Config.Exists("DZTAOTAO_VERSION")=False Then
		dztaotao_Config.Write "DZTAOTAO_VERSION","1.0"
		dztaotao_Config.Write "DZTAOTAO_TITLE_VALUE","大猪滔滔"
		dztaotao_Config.Write "DZTAOTAO_RELEASE_VALUE","5"
		dztaotao_Config.Write "DZTAOTAO_PAGECOUNT_VALUE","12"
		dztaotao_Config.Write "DZTAOTAO_PAGEWIDTH_VALUE","660"
		dztaotao_Config.Write "DZTAOTAO_CHK_VALUE","4"
		dztaotao_Config.Write "DZTAOTAO_CMTCHK_VALUE","4"
		dztaotao_Config.Write "DZTAOTAO_CMTLIMIT_VALUE","999"
		dztaotao_Config.Write "DZTAOTAO_ISIMG_VALUE","1"

		dztaotao_Config.Save
		Call SetBlogHint_Custom("您是第一次安装dztaotao，已经为您导入初始配置。")
	End If

end function



Function dztaotao_Initialize()
	Set dztaotao_Config = New TConfig
	dztaotao_Config.Load("dztaotao")
	DZTAOTAO_TITLE_VALUE = dztaotao_Config.Read ("DZTAOTAO_TITLE_VALUE")
	DZTAOTAO_RELEASE_VALUE=dztaotao_Config.Read ("DZTAOTAO_RELEASE_VALUE")
	DZTAOTAO_PAGECOUNT_VALUE=dztaotao_Config.Read ("DZTAOTAO_PAGECOUNT_VALUE")
	DZTAOTAO_PAGEWIDTH_VALUE=dztaotao_Config.Read ("DZTAOTAO_PAGEWIDTH_VALUE")
	DZTAOTAO_CHK_VALUE=dztaotao_Config.Read ("DZTAOTAO_CHK_VALUE")
	DZTAOTAO_CMTCHK_VALUE=dztaotao_Config.Read ("DZTAOTAO_CMTCHK_VALUE")
	DZTAOTAO_CMTLIMIT_VALUE=dztaotao_Config.Read ("DZTAOTAO_CMTLIMIT_VALUE")
	DZTAOTAO_ISIMG_VALUE=dztaotao_Config.Read("DZTAOTAO_ISIMG_VALUE")
	
End Function


'*********************************************************
' 目的：    
'*********************************************************
Function CheckUpdateDB(a,b)
	Err.Clear
	On Error Resume Next
	Dim Rs
	Set Rs=objConn.execute("SELECT "&a&" FROM "&b)
	Set Rs=Nothing
	If Err.Number=0 Then
	CheckUpdateDB=True
	Else
	Err.Clear
	CheckUpdateDB=False
	End If	
End Function
'*********************************************************

'插入数据表
function dztaotao_CreateDB()

	If Not CheckUpdateDB("[id]","[dz_taotao]") Then
		objConn.execute("CREATE TABLE [dz_taotao] (id AutoIncrement primary key,username VARCHAR(50),site VARCHAR(250),content text,addtime TIME DEFAULT Now(),ttop int DEFAULT 0, tread int DEFAULT 0,comments int DEFAULT 0,itype int DEFAULT 0,img VARCHAR(250),s_img VARCHAR(250))")
		
		objConn.execute("insert into dz_taotao (username,site,content) values ('大猪','http://www.izhu.org','请保证最少保留一条滔滔信息，否则可能出现奇怪的问题。')")
	End If

	If Not CheckUpdateDB("[id]","[dz_comment]") Then
		objConn.execute("CREATE TABLE [dz_comment] (id AutoIncrement primary key,tt_id int DEFAULT 0,u_sername VARCHAR(50),u_site VARCHAR(250),content text,addtime TIME DEFAULT Now(),itype int DEFAULT 0)")
	End If
end function


'安装插件
Function InstallPlugin_dztaotao()

	Call GetFunction()
	Functions(FunctionMetas.GetValue("navbar")).Content=Functions(FunctionMetas.GetValue("navbar")).Content & "<li><a href=""<#ZC_BLOG_HOST#>zb_users/plugin/dztaotao/index.asp"">滔滔</a></li>"
	Functions(FunctionMetas.GetValue("navbar")).Save
	
	Call Activeplugin_dztaotao()
	Call dztaotao_Addto_Navbar()

	Call dztaotao_UpdateConfig()
	Call dztaotao_CreateDB()
	
End Function

'卸载插件
Function UnInstallPlugin_dztaotao()
'On Error Resume Next

	'If CheckUpdateDB("[id]","[dz_taotao]") Then
	'	objConn.execute("DROP TABLE [dz_taotao] ")
	'End If

	'If CheckUpdateDB("[id]","[dz_comment]") Then
	'	objConn.execute("DROP TABLE [dz_comment] ")
	'End If
	Call GetFunction()
	Functions(FunctionMetas.GetValue("navbar")).Content=RemoveLibyUrl(Functions(FunctionMetas.GetValue("navbar")).Content,"<#ZC_BLOG_HOST#>zb_users/plugin/dztaotao/index.asp")
	Functions(FunctionMetas.GetValue("navbar")).Save

	Call dztaotao_Delfrom_Navbar()
Err.Clear
End Function


'添加到导航栏和后台菜单
Function dztaotao_Addto_Navbar()
    Dim strContent, strContent1
    strContent = LoadFromFile(BlogPath & "/include/navbar.asp", "utf-8")
    strContent = strContent & Chr(13)&Chr(10)& "<li><a href='"&ZC_BLOG_HOST&"plugin/dztaotao/'>淘淘</a></li>"
    Call SaveToFile(BlogPath & "/include/navbar.asp", strContent, "utf-8", TRUE)

    strContent1 = LoadFromFile(BlogPath & "admin/admin_left.asp", "utf-8")
    strContent1 = Replace(strContent1, "<p class=""button1""><a onclick='return changeButtonColor(this)' href=""../cmd.asp?act=CategoryMng", "<p class=""button1""><a onclick='return changeButtonColor(this)' href=""../zb_users/plugin/dztaotao/admin.asp?a=list&page=1"" target=""main"">淘淘管理</a></p>"&Chr(13)&Chr(10)&"<p class=""button1""><a onclick='return changeButtonColor(this)' href=""../cmd.asp?act=CategoryMng")
    Call SaveToFile(BlogPath & "admin/admin_left.asp", strContent1, "utf-8", TRUE)
End Function

'从后台菜单和导航栏删除
Function dztaotao_Delfrom_Navbar()
    Dim strContent, strContent1
    strContent = LoadFromFile(BlogPath & "/zb_users/include/navbar.asp", "utf-8")
    strContent = Replace(strContent, Chr(13)&Chr(10)& "<li><a href='"&ZC_BLOG_HOST&"plugin/dztaotao/'>淘淘</a></li>", "")
    Call SaveToFile(BlogPath & "/zb_users/include/navbar.asp", strContent, "utf-8", TRUE)

    strContent1 = LoadFromFile(BlogPath & "admin/admin_left.asp", "utf-8")
    strContent1 = Replace(strContent1, "<p class=""button1""><a onclick='return changeButtonColor(this)' href=""../zb_users/plugin/dztaotao/admin.asp?a=list&page=1"" target=""main"">淘淘管理</a></p>"&Chr(13)&Chr(10)&"<p class=""button1""><a onclick='return changeButtonColor(this)' href=""../cmd.asp?act=CategoryMng", "<p class=""button1""><a onclick='return changeButtonColor(this)' href=""../cmd.asp?act=CategoryMng")
    Call SaveToFile(BlogPath & "admin/admin_left.asp", strContent1, "utf-8", TRUE)
End Function

%>