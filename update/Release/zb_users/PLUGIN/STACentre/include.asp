<%

'注册插件
Call RegisterPlugin("STACentre","ActivePlugin_STACentre")

'具体的接口挂接
Function ActivePlugin_STACentre() 

	'网站管理加上二级菜单项
	Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu",MakeSubMenu("静态管理中心",GetCurrentHost() & "zb_users/plugin/stacentre/main.asp","m-left",False))

End Function
%>