<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("api","ActivePlugin_api")
'挂口部分
Function ActivePlugin_api()
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load("api")
	If objConfig.Exists("version")=False Then
		objConfig.Write "version","1"
		objConfig.Write "id","0"
		objConfig.Write "secret","0"
		objConfig.Write "use_ver","1"
		objConfig.Save
	End If
	
	Call Add_Action_Plugin("Action_Plugin_SiteInfo_Begin",Add_Response_Plugin("Response_Plugin_Admin_SiteInfo","您的API地址为：" & BlogHost & "zb_users/plugin/api/v1/index.asp"))
	 
End Function


Function InstallPlugin_api()

	'用户激活插件之后的操作
	
End Function


Function UnInstallPlugin_api()

	'用户停用插件之后的操作
	
End Function
%>