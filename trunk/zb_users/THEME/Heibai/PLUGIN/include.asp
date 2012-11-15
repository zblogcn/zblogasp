<!-- #include file="Function.asp" -->
<%
'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************
'注册插件
Call RegisterPlugin("Heibai","ActivePlugin_Heibai")
'挂口部分
Function ActivePlugin_Heibai()
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load("Heibai")
	If objConfig.Exists("Version")=False Then
		objConfig.Write "Version","0.2"
		objConfig.Write "SetNewArt","10"
		objConfig.Write "SetCommArt","10"
		objConfig.Write "SetRandomArt","10"
		objConfig.Write "SetNewComm","10"
		objConfig.Write "SetHotCommer","15"
		objConfig.Write "SetTags","30"
		objConfig.Write "SetWeiboSina","http://weibo.com/810888188"
		objConfig.Write "SetWeiboQQ","http://t.qq.com/involvements"
		objConfig.Save
		
		Call SetBlogHint_Custom("<span style='color:#ff0000'>Heibai主题</span>已经激活，点击<a href='" +BlogHost+"ZB_USERS/theme/Heibai/plugin/main.asp'>[主题设置]</a>去配置主题")
	End If
	Set objConfig=Nothing
		
		Call Add_Action_Plugin("Action_Plugin_MakeBlogReBuild_Core_End","CheckArticle")

		Call Add_Filter_Plugin("Filter_Plugin_TArticleList_Build_Template_Succeed","HeibaiTConfig")
		Call Add_Filter_Plugin("Filter_Plugin_TArticle_Build_Template_Succeed","HeibaiTConfig")

		
		Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(1,"主题设置",BlogHost&"ZB_USERS/theme/Heibai/plugin/main.asp","aHeibai",""))'添加右上角导航
End Function

Function HeibaiTConfig(html)
	Dim objConfig,ZC_TM_SetWeiboSina,ZC_TM_SetWeiboQQ
	Set objConfig=New TConfig
	objConfig.Load("Heibai")
	ZC_TM_SetWeiboSina=objConfig.Read("SetWeiboSina")
	ZC_TM_SetWeiboQQ=objConfig.Read("SetWeiboQQ")
	html=Replace(html,"<#ZC_TM_SetWeiboSina#>",ZC_TM_SetWeiboSina)
	html=Replace(html,"<#ZC_TM_SetWeiboQQ#>",ZC_TM_SetWeiboQQ)
	Set objConfig=Nothing
End Function

'检查所有列表===================================
Function CheckArticle()
	Call CheckNewArticle()
	Call CheckCommArticle()
	Call CheckRandomArticle()

	Call CheckRandomArticle()
	Call CheckNewComm()
	Call CheckHotCommer()

End Function

'卸载所有列表===================================
Function RemArticle()
	Call RemNewArticle()
	Call RemCommArticle()
	Call RemRandomArticle()
End Function

Function RemCom()
	Call RemNewComm()
	Call RemHotCommer()
End Function

'================================操作==========================
Function RemHeibaiTConfig()
	Dim objConfig
	Set objConfig=New TConfig
	objConfig.Load("Heibai")
	If objConfig.Exists("Version")=True Then
		objConfig.Delete
	End If
	Set objConfig=Nothing
End Function

'安装插件
Function InstallPlugin_Heibai
	Call SetBlogHint(Empty,Empty,True)
End Function

'卸载插件
Function UnInstallPlugin_Heibai
	Call RemArticle()
	Call RemCom()
	Call SetBlogHint(Empty,Empty,True)
	Call RemHeibaiTConfig()
End Function
%>