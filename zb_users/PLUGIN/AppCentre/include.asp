<%


Call RegisterPlugin("AppCentre","ActivePlugin_AppCentre")
'挂口部分
Function ActivePlugin_AppCentre()

	'Call Add_Response_Plugin("Response_Plugin_PluginMng_SubMenu",MakeSubMenu("资源中心",BlogHost & "zb_users/plugin/AppCentre/main.asp?act=p","m-left",False))

	'Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu",MakeSubMenu("资源中心",BlogHost & "zb_users/plugin/AppCentre/main.asp?act=t","m-left",False))


	Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu",MakeSubMenu("在线安装主题<script src='"& BlogHost &"zb_users/plugin/appcentre/theme.js' type='text/javascript'></script>",BlogHost & "zb_users/plugin/appcentre/server.asp?","m-left",False))

	Call Add_Response_Plugin("Response_Plugin_PluginMng_SubMenu",MakeSubMenu("在线安装插件<script src='"& BlogHost &"zb_users/plugin/appcentre/plugin.js' type='text/javascript'></script>",BlogHost & "zb_users/plugin/appcentre/server.asp?","m-left",False))

	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(1,"应用中心",GetCurrentHost&"zb_users/plugin/appcentre/main.asp","nav_appcentre","aAppcentre",GetCurrentHost&"zb_users/plugin/appcentre/images/web.png"))
	
	'检查更新
	Call Add_Response_Plugin("Response_Plugin_SiteInfo_SubMenu","<script type='text/javascript'>$.get(bloghost+'zb_users/plugin/appcentre/checkupdate.asp?act=silent');</script>")

	Call Add_Action_Plugin("Action_Plugin_Admin_Begin","Call AppCentre_NewVersionFound()")

End Function


Function AppCentre_NewVersionFound()
	Dim o
	o=LoadFromFile(BlogPath&"zb_users\cache\appcentre_list.lst","utf-8")
	If Replace(o,",","")<>"" Then
		Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(1,"<font color='red'>发现插件更新</font>",BlogHost&"zb_users/plugin/appcentre/checkupdate.asp","AppCentre_Update",""))
	End If
End Function
%>