<!-- #include file="functions.asp"-->
<%
'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************
'注册插件
Call RegisterPlugin("AdvancedFunction","ActivePlugin_AdvancedFunction")
'挂口部分
Call Add_Response_Plugin("Response_Plugin_Html_Js_Add","$(document).ready(function(){$.getScript(bloghost+'zb_users/plugin/advancedfunction/random.asp')});")

Const JSEMPTY=Empty

Function ActivePlugin_AdvancedFunction()
	Dim aa,bb,cc
	aa="随机文章,访问最多文章,本月最热文章,本年最热文章,分类最热文章,分类"
	bb="评论最多文章,本月评论最多,本年评论最多,分类评论最多"
	cc=aa&","&bb
	Call Add_Action_Plugin("Action_Plugin_MakeBlogReBuild_Core_Begin","advancedfunction.run("""&cc&""")")
	
End Function
%>