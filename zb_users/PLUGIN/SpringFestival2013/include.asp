<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("SpringFestival2013","ActivePlugin_SpringFestival2013")
'挂口部分
Function ActivePlugin_SpringFestival2013()

	Call Add_Response_Plugin("Response_Plugin_Admin_Header","<style type=""text/css"">.top{background:url(" & BlogHost & "zb_users/plugin/springfestival2013/top.jpg) !important}</style>")
	
End Function


Function InstallPlugin_SpringFestival2013()

	'用户激活插件之后的操作
	
End Function


Function UnInstallPlugin_SpringFestival2013()

	'用户停用插件之后的操作
	
End Function
%>