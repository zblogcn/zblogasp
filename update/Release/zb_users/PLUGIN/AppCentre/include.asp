<%


Call RegisterPlugin("AppCentre","ActivePlugin_AppCentre")
'挂口部分
Function ActivePlugin_AppCentre()

	Dim ac
	Set ac=New TConfig
	ac.Load "AppCentre"

	If BlogUser.Level=1 Then Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu",MakeSubMenu("在线安装主题<script src='"& BlogHost &"zb_users/plugin/appcentre/theme_js.asp' type='text/javascript'></script>",BlogHost & "zb_users/plugin/appcentre/server.asp?cate=2","m-left",False))

	If BlogUser.Level=1 Then Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu",MakeSubMenu("编辑当前主题信息",BlogHost & "zb_users/plugin/appcentre/theme_edit.asp?id="&Server.URLEncode(ZC_BLOG_THEME),"m-left",False))

	If BlogUser.Level=1 Then Call Add_Response_Plugin("Response_Plugin_PluginMng_SubMenu",MakeSubMenu("在线安装插件<script src='"& BlogHost &"zb_users/plugin/appcentre/plugin_js.asp' type='text/javascript'></script>",BlogHost & "zb_users/plugin/appcentre/server.asp?cate=1","m-left",False))

	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(1,"应用中心",GetCurrentHost&"zb_users/plugin/appcentre/main.asp","nav_appcentre","aAppcentre",BlogHost&"zb_users/plugin/appcentre/images/cube1.png"))

	'检查更新
	If BlogUser.Level=1 Then
		Dim last
		last=ac.read("LastChechUpdate")
		If last="" Then last="2000-01-01"
		last=Replace(last,"|","")
		If DateDiff("h", last, Now)>=11 Then
			Randomize
			Call Add_Response_Plugin("Response_Plugin_SiteInfo_SubMenu","<script type='text/javascript'>$(document).ready(function(){  $.getScript('"&BlogHost&"zb_users/plugin/appcentre/server.asp?method=checksilent&rnd="&Rnd()&"'); });</script>")
			ac.Write "LastChechUpdate","|"&Now&"|"
			ac.Save			
		End If
	End If

	Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu",MakeSubMenu("应用中心设置",GetCurrentHost() & "zb_users/plugin/appcentre/setting.asp","m-left",False))

End Function
%>