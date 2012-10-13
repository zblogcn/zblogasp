<%

'注册插件
Call RegisterPlugin("ZBDK","ActivePlugin_ZBDK")
'挂口部分
Function ActivePlugin_ZBDK()

	Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(1,"开发工具",BlogHost&"zb_users/plugin/ZBDK/main.asp","zbdk",""))
	
End Function
%>