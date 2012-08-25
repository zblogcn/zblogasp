<%


Call RegisterPlugin("AppCentre","ActivePlugin_AppCentre")
'挂口部分
Function ActivePlugin_AppCentre()

	Call Add_Response_Plugin("Response_Plugin_PluginMng_SubMenu",MakeSubMenu("资源中心",BlogHost & "zb_users/plugin/AppCentre/main.asp?act=p","m-left",False))

	Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu",MakeSubMenu("资源中心",BlogHost & "zb_users/plugin/AppCentre/main.asp?act=t","m-left",False))
End Function

%>