<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("weixin_Search","ActivePlugin_weixin_Search")
'挂口部分
Function ActivePlugin_weixin_Search()

	'插件最主要在这里挂接口。
	'Z-Blog可挂的接口有三类：Action、Filter、Response
	'建议参考Z-Wiki进行开发
	
End Function


Function InstallPlugin_weixin_Search()

	'用户激活插件之后的操作
	
End Function


Function UnInstallPlugin_weixin_Search()

	'用户停用插件之后的操作
	
End Function
%>