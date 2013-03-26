<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("AdvancedNavbar","ActivePlugin_AdvancedNavbar")
'挂口部分
Function ActivePlugin_AdvancedNavbar()

	'插件最主要在这里挂接口。
	'Z-Blog可挂的接口有三类：Action、Filter、Response
	'建议参考Z-Wiki进行开发
	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(1,"高级导航栏",GetCurrentHost&"zb_users/plugin/AdvancedNavbar/main.asp","nav_quoted","aAdvancedNavbar",""))
	Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu",MakeSubMenu("可视化导航栏设置",GetCurrentHost() & "zb_users/plugin/AdvancedNavbar/main.asp","m-left",False))
End Function


Function InstallPlugin_AdvancedNavbar()
	'用户激活插件之后的操作
	InstallAdvancedNavbar()
End Function


Function UnInstallPlugin_AdvancedNavbar()
	'用户停用插件之后的操作
	UnInstallAdvancedNavbar()
End Function
%>