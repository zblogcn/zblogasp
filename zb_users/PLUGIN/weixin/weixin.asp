<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("weixin","ActivePlugin_weixin")
'挂口部分
Function ActivePlugin_weixin()

	'插件最主要在这里挂接口。
	'Z-Blog可挂的接口有三类：Action、Filter、Response
	'建议参考Z-Wiki进行开发
	
End Function


Function InstallPlugin_weixin()

	'用户激活插件之后的操作
	
End Function


Function UnInstallPlugin_weixin()

	'用户停用插件之后的操作
	
End Function
%>