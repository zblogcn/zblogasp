<%

'注册插件
Call RegisterPlugin("AdvancedConfig","ActivePlugin_AdvancedConfig")
'挂口部分
Function ActivePlugin_AdvancedConfig()

	'Call Add_Action_Plugin("Action_Plugin_Edit_Setting_Begin","Response.Redirect BlogHost & ""/zb_users/plugin/advancedconfig/main.asp""")
	'网站管理加上二级菜单项
	Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu",MakeSubMenu("高级设置",GetCurrentHost() & "zb_users/plugin/advancedconfig/main.asp","m-right",False))
End Function
%>