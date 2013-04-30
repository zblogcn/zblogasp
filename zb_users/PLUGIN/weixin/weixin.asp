<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("weixin","ActivePlugin_weixin")
'挂口部分
Function ActivePlugin_weixin()
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load("weixin")
	If objConfig.Exists("Version")=False Then
		objConfig.Write "Version","1.0"
		objConfig.Write "WelcomeStr","欢迎关注《{%title%}》！！！<br/>您可发送“最新文章”来查看博客最新的{%num%}篇文章，或者直接发送关键词来搜索博客中已发表的文章。更多使用帮助请输入英文“help”来查看。"
		objConfig.Write "SearchNum","10"
		objConfig.Write "LastPostNum","5"
		objConfig.Write "ShowMeta","1"
		objConfig.Write "token","weixin"
		objConfig.Save
	End If
End Function


Function InstallPlugin_weixin()
	'用户激活插件之后的操作
End Function


Function UnInstallPlugin_weixin()
	'用户停用插件之后的操作	
End Function
%>