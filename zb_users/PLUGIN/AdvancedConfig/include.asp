<%

'注册插件
Call RegisterPlugin("AdvancedConfig","ActivePlugin_AdvancedConfig")
'挂口部分
Function ActivePlugin_AdvancedConfig()

	Call Add_Action_Plugin("Action_Plugin_Edit_Setting_Begin","Response.Redirect BlogHost & ""/zb_users/plugin/advancedconfig/main.asp""")
	
End Function
%>