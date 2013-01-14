<%


Call RegisterPlugin("AppCentre","ActivePlugin_AppCentre")
'挂口部分
Function ActivePlugin_AppCentre()

	Dim ac
	Set ac=New TConfig
	ac.Load "AppCentre"

	If BlogUser.Level=1 Then Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu",MakeSubMenu("在线安装主题<script type='text/javascript'>var disableupdatetheme="""&ac.read("DisableUpdateTheme")&""";</script><script src='"& BlogHost &"zb_users/plugin/appcentre/theme_js.asp' type='text/javascript'></script>",BlogHost & "zb_users/plugin/appcentre/server.asp?action=catalog&cate=2","m-left",False))

	If BlogUser.Level=1 Then Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu",MakeSubMenu("编辑当前主题信息",BlogHost & "zb_users/plugin/appcentre/theme_edit.asp?id="&Server.URLEncode(ZC_BLOG_THEME),"m-left",False))

	If BlogUser.Level=1 Then Call Add_Response_Plugin("Response_Plugin_PluginMng_SubMenu",MakeSubMenu("在线安装插件<script src='"& BlogHost &"zb_users/plugin/appcentre/plugin_js.asp' type='text/javascript'></script>",BlogHost & "zb_users/plugin/appcentre/server.asp?action=catalog&cate=1","m-left",False))

	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(1,"应用中心",GetCurrentHost&"zb_users/plugin/appcentre/main.asp","nav_appcentre","aAppcentre",BlogHost&"zb_users/plugin/appcentre/images/cube1.png"))

	'检查更新
	If BlogUser.Level=1 Then
		Randomize
		Call Add_Response_Plugin("Response_Plugin_SiteInfo_SubMenu","<script type='text/javascript'>$.get(bloghost+'zb_users/plugin/appcentre/server.asp?action=update&silent=true&rnd="&Rnd()&"',function(data){eval(data);});</script>")
	End If

End Function
%>