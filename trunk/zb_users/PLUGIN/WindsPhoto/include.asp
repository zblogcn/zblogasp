<%
'注册插件
Dim WP_DATA_PATH,WP_ALBUM_NAME,WP_ALBUM_INTRO,WP_SUB_DOMAIN,WP_SCRIPT_TYPE,WP_ORDER_BY,WP_SMALL_WIDTH,WP_SMALL_HEIGHT,WP_LIST_WIDTH,WP_LIST_HEIGHT,WP_UPLOAD_FILESIZE,WP_UPLOAD_DIR,WP_UPLOAD_DIRBY,WP_UPLOAD_RENAME,WP_WATERMARK_WIDTH_POSITION,WP_WATERMARK_HEIGHT_POSITION,WP_JPEG_FONTCOLOR,WP_JPEG_FONTBOLD,WP_JPEG_FONTSIZE,WP_JPEG_FONTQUALITY,WP_WATERMARK_AUTO,WP_WATERMARK_TYPE,WP_WATERMARK_TEXT,WP_WATERMARK_LOGO,WP_WATERMARK_ALPHA,WP_INDEX_PAGERCOUNT,WP_SMALL_PAGERCOUNT,WP_LIST_PAGERCOUNT,WP_BLOGPHOTO_ID,WP_IF_ASPJPEG,WP_HIDE_DIVFILESND

Call Registerplugin("WindsPhoto", "Activeplugin_WindsPhoto")
Dim WP_Config
Function WindsPhoto_Initialize()
	Set WP_Config=New TConfig
	WP_Config.Load "WindsPhoto"
	If WP_Config.Exists("WP_VER")=False Then
		WP_Config.Write "WP_VER","2.7.4":WP_Config.Write "WP_DATA_PATH","data/windsphoto.mdb":WP_Config.Write "WP_ALBUM_NAME","我的WindsPhoto":WP_Config.Write "WP_ALBUM_INTRO","<p>WindsPhoto是基于asp+access的Z-Blog图片相册管理插件，功能简介实用。</p>":WP_Config.Write "WP_SUB_DOMAIN","":WP_Config.Write "WP_SCRIPT_TYPE","4":WP_Config.Write "WP_ORDER_BY","1":WP_Config.Write "WP_SMALL_WIDTH",144:WP_Config.Write "WP_SMALL_HEIGHT",144:WP_Config.Write "WP_LIST_WIDTH",600:WP_Config.Write "WP_LIST_HEIGHT",600:WP_Config.Write "WP_UPLOAD_FILESIZE",2048000:WP_Config.Write "WP_UPLOAD_DIR","photofile":WP_Config.Write "WP_UPLOAD_DIRBY","1":WP_Config.Write "WP_UPLOAD_RENAME","1":WP_Config.Write "WP_WATERMARK_WIDTH_POSITION","right":WP_Config.Write "WP_WATERMARK_HEIGHT_POSITION","bottom":WP_Config.Write "WP_JPEG_FONTCOLOR","#000":WP_Config.Write "WP_JPEG_FONTBOLD","true":WP_Config.Write "WP_JPEG_FONTSIZE","14":WP_Config.Write "WP_JPEG_FONTQUALITY","4":WP_Config.Write "WP_WATERMARK_AUTO","0":WP_Config.Write "WP_WATERMARK_TYPE","1":WP_Config.Write "WP_WATERMARK_TEXT","WindsPhoto":WP_Config.Write "WP_WATERMARK_LOGO","images/nopic.jpg":WP_Config.Write "WP_WATERMARK_ALPHA","0.7":WP_Config.Write "WP_INDEX_PAGERCOUNT",12:WP_Config.Write "WP_SMALL_PAGERCOUNT",18:WP_Config.Write "WP_LIST_PAGERCOUNT",8:WP_Config.Write "WP_BLOGPHOTO_ID",1:WP_Config.Write "WP_IF_ASPJPEG","1":WP_Config.Write "WP_HIDE_DIVFILESND","1":WP_Config.Save
	End If
	WP_DATA_PATH = WP_Config.Read("WP_DATA_PATH")
	WP_ALBUM_NAME = WP_Config.Read("WP_ALBUM_NAME")
	WP_ALBUM_INTRO = WP_Config.Read("WP_ALBUM_INTRO")
	WP_SUB_DOMAIN = WP_Config.Read("WP_SUB_DOMAIN")
	WP_SCRIPT_TYPE = WP_Config.Read("WP_SCRIPT_TYPE")
	WP_ORDER_BY = WP_Config.Read("WP_ORDER_BY")
	WP_SMALL_WIDTH = CInt(WP_Config.Read("WP_SMALL_WIDTH"))
	WP_SMALL_HEIGHT = CInt(WP_Config.Read("WP_SMALL_HEIGHT"))
	WP_LIST_WIDTH = CInt(WP_Config.Read("WP_LIST_WIDTH"))
	WP_LIST_HEIGHT = CInt(WP_Config.Read("WP_LIST_HEIGHT"))
	WP_UPLOAD_FILESIZE = CLng(WP_Config.Read("WP_UPLOAD_FILESIZE"))
	WP_UPLOAD_DIR = WP_Config.Read("WP_UPLOAD_DIR")
	WP_UPLOAD_DIRBY = WP_Config.Read("WP_UPLOAD_DIRBY")
	WP_UPLOAD_RENAME = WP_Config.Read("WP_UPLOAD_RENAME")
	WP_WATERMARK_WIDTH_POSITION = WP_Config.Read("WP_WATERMARK_WIDTH_POSITION")
	WP_WATERMARK_HEIGHT_POSITION = WP_Config.Read("WP_WATERMARK_HEIGHT_POSITION")
	WP_JPEG_FONTCOLOR = WP_Config.Read("WP_JPEG_FONTCOLOR")
	WP_JPEG_FONTBOLD = WP_Config.Read("WP_JPEG_FONTBOLD")
	WP_JPEG_FONTSIZE = WP_Config.Read("WP_JPEG_FONTSIZE")
	WP_JPEG_FONTQUALITY = WP_Config.Read("WP_JPEG_FONTQUALITY")
	WP_WATERMARK_AUTO = WP_Config.Read("WP_WATERMARK_AUTO")
	WP_WATERMARK_TYPE = WP_Config.Read("WP_WATERMARK_TYPE")
	WP_WATERMARK_TEXT = WP_Config.Read("WP_WATERMARK_TEXT")
	WP_WATERMARK_LOGO = WP_Config.Read("WP_WATERMARK_LOGO")
	WP_WATERMARK_ALPHA = WP_Config.Read("WP_WATERMARK_ALPHA")
	WP_INDEX_PAGERCOUNT = CInt(WP_Config.Read("WP_INDEX_PAGERCOUNT"))
	WP_SMALL_PAGERCOUNT = CInt(WP_Config.Read("WP_SMALL_PAGERCOUNT"))
	WP_LIST_PAGERCOUNT = CInt(WP_Config.Read("WP_LIST_PAGERCOUNT"))
	WP_BLOGPHOTO_ID = CInt(WP_Config.Read("WP_BLOGPHOTO_ID"))
	WP_IF_ASPJPEG = WP_Config.Read("WP_IF_ASPJPEG")
	WP_HIDE_DIVFILESND = WP_Config.Read("WP_HIDE_DIVFILESND")
End Function

'安装插件
Function Installplugin_WindsPhoto()
    On Error Resume Next
    Call WindsPhoto_Addto_Navbar()
    Call WindsPhoto_Copy_Template()
    Call WindsPhoto_Database_Rename()
    Call SetBlogHint_Custom("? 提示:[WindsPhoto]已启用,现在进入初始化系统设置.升级请手动修改include.asp数据库参数.2.7以下版本请<a href=""update.asp"">升级数据库结构</a>")
    Response.Redirect BlogHost &"zb_users/plugin/WindsPhoto/admin_setting.asp"
    Err.Clear
End Function

'卸载插件
Function UnInstallplugin_WindsPhoto()
    Call DelSiteFile("/zb_users/include/windsphoto_sort.asp")
    Call DelSiteFile("/photo.html")
    Call WindsPhoto_Delfrom_Navbar
    Call SetBlogHint_Custom("? 提示:[WindsPhoto]已停用,由本插件生成的文件已删除,你上传的图片文件以及数据库仍然保留,卸载请手动删除插件目录.")
End Function

Function Activeplugin_WindsPhoto()
    '网站管理加上二级菜单项
    Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu", MakeSubMenu("WindsPhoto设置", BlogHost & "zb_users/plugin/windsphoto/admin_setting.asp", "m-left", FALSE))
    Call Add_Response_Plugin("Response_Plugin_SiteInfo_SubMenu",MakeSubMenu("[WindsPhoto管理]",BlogHost & "zb_users/plugin/windsphoto/admin_main.asp","m-left",False))
	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(2,"相册管理",GetCurrentHost&"zb_users/plugin/windsphoto/admin_main.asp","nav_windsphoto","aWindsPhoto",GetCurrentHost&"zb_users/plugin/windsphoto/images/MD-camera-photo.png"))
    If WP_BLOGPHOTO_ID <> 0 then
    '    Call Add_Action_Plugin("Action_Plugin_Edit_Begin","Call WindsPhoto_addForm()")
     '   Call Add_Action_Plugin("Action_Plugin_Edit_Fckeditor_Begin","Call WindsPhoto_addForm()")
    end If
End Function


'首次安装数据库改名....
Function WindsPhoto_Database_Rename()
	Call WindsPhoto_Initialize
    Dim fso, f, s, pathnew, ranNum
    Randomize
    ranNum = Int((99 -10 + 1) * Rnd + 99)
    pathnew = Year(Now)&Month(Now)&Day(Now)&Hour(Now)&Minute(Now)&Second(Now)&ranNum& ".mdb"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(BlogPath & "/zb_users/plugin/windsphoto/data/windsphoto.mdb") Then
        Set f = fso.GetFile(BlogPath & "/zb_users/plugin/windsphoto/"& WP_DATA_PATH)
        f.Name = pathnew
        Set f = Nothing
        Dim strContent, strWP_DATA_PATH
        strWP_DATA_PATH = "data/" & pathnew
        WP_Config.Write "WP_DATA_PATH", strWP_DATA_PATH
		WP_Config.Save
    End If
End Function

'添加到导航栏和后台菜单
Function WindsPhoto_Addto_Navbar()
	Call GetFunction()
	Functions(FunctionMetas.GetValue("navbar")).Content=Functions(FunctionMetas.GetValue("navbar")).Content & "<li><a href=""<#ZC_BLOG_HOST#>zb_users/plugin/windsphoto/"">Photos</a></li>"
	Functions(FunctionMetas.GetValue("navbar")).Save
	Call ClearGlobeCache
	Call LoadGlobeCache
End Function

'从后台菜单和导航栏删除
Function WindsPhoto_Delfrom_Navbar()
	Call GetFunction()
	Functions(FunctionMetas.GetValue("navbar")).Content=RemoveLibyUrl(Functions(FunctionMetas.GetValue("navbar")).Content,"<#ZC_BLOG_HOST#>zb_users/plugin/windsphoto/")
	Functions(FunctionMetas.GetValue("navbar")).Save
	Call ClearGlobeCache
	Call LoadGlobeCache
End Function

'添加相册默认模板
Function WindsPhoto_Copy_Template()
    Dim strContent
    strContent = LoadFromFile(BlogPath & "/zb_users/Themes/" & ZC_BLOG_THEME & "/Template/tags.html", "utf-8")
    strContent = Replace(strContent, "<#ZC_MSG050#>", "相册分类")
    strContent = Replace(strContent, "<#CACHE_INCLUDE_CALENDAR#>", "<ul><#CACHE_INCLUDE_WINDSPHOTO_SORT#></ul>")
    strContent = Replace(strContent, "id=""divCalendar""", "id=""divCatalog""")
    Call SaveToFile(BlogPath & "/zb_users/Themes/" & ZC_BLOG_THEME & "/Template/wp_index.html", strContent, "utf-8", TRUE)
    Call SaveToFile(BlogPath & "/zb_users/Themes/" & ZC_BLOG_THEME & "/Template/wp_album.html", strContent, "utf-8", TRUE)
End Function
%>