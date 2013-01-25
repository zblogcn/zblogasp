<%
	Call ZBDK_AddStatus("OnlinePlugin","在线插件",BlogHost & "zb_users/plugin/ZBDK/OnlinePlugin/main.asp","不需要再创建或编辑现有插件就可以挂接口的工具（呃。。）")
	Call Add_Action_Plugin("Action_Plugin_ZBDK_Else","ZBDK_Else")
%>