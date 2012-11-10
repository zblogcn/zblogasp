<!-- #include file="functions.asp"-->
<%
Const jempty=Empty
'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************
'注册插件
Call RegisterPlugin("AdvancedFunction","ActivePlugin_AdvancedFunction")
'挂口部分
Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","$(document).ready(function(){$.getScript(bloghost+'zb_users/plugin/advancedfunction/random.asp')})")

Function ActivePlugin_AdvancedFunction()
	Dim aa,bb,cc
	aa="随机文章,访问最多文章,本月最热文章,本年最热文章,分类最热文章,分类"
	bb="评论最多文章,本月评论最多,本年评论最多,分类评论最多"
	cc=aa&","&bb
	Call Add_Action_Plugin("Action_Plugin_MakeBlogReBuild_Core_Begin","advancedfunction.run("""&cc&""")")
	
End Function

Function AdvancedFunction_Initialize

End Function

'检查所有阅读列表===================================
Function CheckViewAll()
	Call CheckHotArticle()
	Call CheckHotMArticle()
	Call CheckHotYArticle()
	Call CheckRandomArticle()
End Function

'检查所有列表===================================
Function CheckArticle()
	Call CheckNewArticle()
	Call CheckHotArticle()
	Call CheckCommArticle()
	Call CheckHotMArticle()
	Call CheckHotYArticle()
	Call CheckRandomArticle()
	Call CheckCategroyForNew()
	Call CheckCom()
End Function

Function CheckCom()
	Call CheckHotComm()
	Call CheckHotGuestbook()
End Function

'卸载所有列表===================================
Function RemArticle()
	Call RemHotYArticle()
	Call RemHotMArticle()
	Call RemCommArticle()
	Call RemRandomArticle()
	Call RemHotArticle()
	Call RemNewArticle()
End Function

Function RemCom()
	Call RemHotComm()
	Call RemHotGuestbook()
End Function


'安装插件
Function InstallPlugin_AdvancedFunction

	Call CheckArticle()
	Call CheckCom()
	Call SetBlogHint(Empty,Empty,True)
	
End Function

'卸载插件
Function UnInstallPlugin_AdvancedFunction
	
	Call RemArticle()
	Call RemCom()
	Call SetBlogHint(Empty,Empty,True)

End Function
%>