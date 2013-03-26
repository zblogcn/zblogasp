<!-- #include file="Function.asp" -->
<%
'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************
'注册插件
Call RegisterPlugin("X2013","ActivePlugin_X2013")
'挂口部分
Function ActivePlugin_X2013()
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load("X2013")
	If objConfig.Exists("Version")=False Then
		objConfig.Write "Version","1.0"
		objConfig.Write "SetWeiboSina","http://weibo.com/810888188"
		objConfig.Write "SetWeiboQQ","http://t.qq.com/involvements"
		objConfig.Write "DisplayFeed","True"
		objConfig.Write "SetMailKey","4e54e0008863773ff0f44e54eb9c1805cf165e63a0601789"
		objConfig.Write "PostAdHeader",""
		objConfig.Write "PostAdFooter",""
		objConfig.Save
		Call SetBlogHint_Custom("<span style='color:#ff0000'>X2013主题</span>已经激活，点击<a href='" +BlogHost+"ZB_USERS/theme/X2013/plugin/main.asp'>[主题设置]</a>去配置主题")
	End If
		'Call Add_Action_Plugin("Filter_Plugin_TArticle_Build_Template_Succeed","X2013TConfig")
		Call Add_Action_Plugin("Action_Plugin_TArticleList_Export_Begin","Call Add_Filter_Plugin(""Filter_Plugin_TArticleList_Build_Template"",""X2013TConfig"")")
		Call Add_Action_Plugin("Action_Plugin_TArticle_Export_Begin","Call Add_Filter_Plugin(""Filter_Plugin_TArticle_Build_Template"",""X2013TConfig"")")
		Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(1,"X2013主题设置",BlogHost & "zb_users/theme/X2013/plugin/main.asp","aX2013",""))
End Function

Function X2013TConfig(html)
	Dim objConfig,ZC_TM_SetWeiboSina,ZC_TM_SetWeiboQQ,ZC_TM_SetWeiBo,ZC_TM_SetFeedToMail,SetMailKey,MailFeedhtml,ZC_TM_POSTADHEADER,ZC_TM_POSTADFOOTER
	Set objConfig=New TConfig
	objConfig.Load("X2013")
	'============微博
	ZC_TM_SetWeiboSina=objConfig.Read("SetWeiboSina")
	ZC_TM_SetWeiboQQ=objConfig.Read("SetWeiboQQ")
	If ZC_TM_SetWeiboQQ<>"" Then ZC_TM_SetWeiboQQ="<li><a class=""btn btn-mini"" target=""_blank"" href="""&ZC_TM_SetWeiboQQ&""">腾讯微博</a></li>"
	If ZC_TM_SetWeiboSina<>"" Then ZC_TM_SetWeiboSina="<li><a class=""btn btn-mini"" target=""_blank"" href="""&ZC_TM_SetWeiboSina&""">新浪微博</a></li>"
	ZC_TM_SetWeiBo="<ul class=""popup-follow-weibo"">"&ZC_TM_SetWeiboQQ&ZC_TM_SetWeiboSina&"</ul>"
	If ZC_TM_SetWeiboSina="" and ZC_TM_SetWeiboQQ="" Then ZC_TM_SetWeiBo=""
	html=Replace(html,"<#ZC_TM_SetWeiBo#>",ZC_TM_SetWeiBo)
	'============Feed订阅
	If (CBool(objConfig.Read("DisplayFeed")))=True Then
		SetMailKey=objConfig.Read("SetMailKey")
		ZC_TM_SetFeedToMail="<div class=""popup-follow-mail""><h4>邮件订阅：</h4><form action=""http://list.qq.com/cgi-bin/qf_compose_send"" target=""_blank"" method=""post""><input type=""hidden"" name=""t"" value=""qf_booked_feedback"" /><input type=""hidden"" name=""id"" value="""&SetMailKey&""" /><input id=""to"" placeholder=""输入邮箱 订阅本站"" name=""to"" type=""text"" class=""ipt"" /><input class=""btn btn-primary"" type=""submit"" value=""邮件订阅"" /></form></div>"
		html=Replace(html,"<#ZC_TM_SetFeedToMail#>",ZC_TM_SetFeedToMail)
	Else
		html=Replace(html,"<#ZC_TM_SetFeedToMail#>",ZC_TM_SetFeedToMail)
	End If
	'==========AD
	ZC_TM_POSTADHEADER=objConfig.Read("PostAdHeader")
	ZC_TM_POSTADFOOTER=objConfig.Read("PostAdFooter")
	html=Replace(html,"<#ZC_TM_POSTADHEADER#>",ZC_TM_POSTADHEADER)
	html=Replace(html,"<#ZC_TM_POSTADFOOTER#>",ZC_TM_POSTADFOOTER)
End Function

'================================操作==========================
Function RemX2013TConfig()
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load("X2013")
	If objConfig.Exists("Version")=True Then
		objConfig.Delete
	End If
End Function

'安装插件
Function InstallPlugin_X2013
	Call SetBlogHint(Empty,Empty,True)
End Function

'卸载插件
Function UnInstallPlugin_X2013
	Call SetBlogHint(Empty,Empty,True)
End Function
%>