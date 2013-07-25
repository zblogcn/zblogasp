<!-- #include file="option.asp" -->
<%
'注册插件
Call RegisterPlugin("Wap","ActivePlugin_Wap")
'挂口部分
Function ActivePlugin_Wap()
	Call WAPConfig_Initialize()
	Call Add_Response_Plugin("Response_Plugin_SettingMng_SubMenu",MakeSubMenu("WAP设置",GetCurrentHost() & "zb_users/plugin/wap/main.asp","m-left",False))
End Function

'首页接口，判断类型并跳转
Call Add_Action_Plugin("Action_Plugin_Default_WithOutConnect_Begin","Wap_Check()")
Call Add_Response_Plugin("Response_Plugin_HTML_JS_Add",Wap_AddControl)

Function Wap_AddControl()
	Wap_AddControl="$(document).ready(function(){if(/Android|iOS|Windows Phone|IEMobile|BB10|ADR|iPhone|iPad|iPod|Symbian|BlackBerry/i.test(navigator.userAgent)){$('.cp-vrs').after(""&nbsp;&nbsp;<span class='cp-pad'><a href='""+bloghost+""?mod=pad&act=editarticle'>[新建文章(PAD)]</a></span>"")}});"
End Function 

'检查终端类型
Dim Wap_Type
Function Wap_Check()

	Dim s
	'Response.Write ZC_DISPLAY_COUNT_WAP
	If Request.QueryString("mod")="pc" Then Exit Function

	If Request.QueryString("mod")="pad" Then
		Server.Transfer "zb_users\plugin\wap\pad.asp":Response.End
	End If

	If Request.QueryString("mod")="wap" Then
		If WAP_DISABLE=True Then Server.Transfer "zb_users\plugin\wap\pad.asp":Response.End
		Server.Transfer "zb_users\plugin\wap\wap.asp":Response.End
	End If

	Wap_Type=Wap_CheckMobile()
	If Wap_Type="pad" Then
		Server.Transfer "zb_users\plugin\wap\pad.asp":Response.End
	End If
	If Wap_Type="wap" Then
		If WAP_DISABLE=True Then Server.Transfer "zb_users\plugin\wap\pad.asp":Response.End
		Server.Transfer "zb_users\plugin\wap\wap.asp":Response.End
	End If

End Function 


'检查手机端类型
Function Wap_CheckMobile()

	Wap_CheckMobile=""

	Dim Mobile_List,Pad_List,UserAgent
	UserAgent = LCase(Request.ServerVariables("HTTP_USER_AGENT"))

	Pad_List="android|iphone|ipad|windows\sphone|kindle|gt\-p|gt\-n|meego"
	If CheckRegExp(UserAgent,Pad_List) Then 
		Wap_CheckMobile="pad"
		Exit Function
	End If

	'是否（智能）手机浏览器
	Mobile_List ="android|iphone|ipad|windows\sphone|kindle|rim\stablet|meego|netfront|java|opera\smini|opera\smobi|ucweb|windows\sce|symbian|series|webos|sonyericsson|sony|blackberry|cellphone|dopod|nokia|samsung|palmsource|palmos|xphone|xda|smartphone|meizu|up.browser|up.link|pieplus|midp|cldc|motorola|foma|docomo|huawei|coolpad|alcatel|amoi|ktouch|philips|benq|haier|bird|zte|wap|mobile"
	If CheckRegExp(UserAgent,Mobile_List) Then 
		Wap_CheckMobile="wap"
		Exit Function
	End If

	'是否专用wap浏览器
	If InStr(LCase(Request.ServerVariables("HTTP_ACCEPT")), "application/vnd.wap.xhtml+xml") Then
		Wap_CheckMobile="wap"
		Exit Function
	End If
	If InStr(LCase(Request.ServerVariables("HTTP_VIA")), "wap")>0 Then
		Wap_CheckMobile="wap"
		Exit Function
	End If
	If Not IsEmpty(Request.ServerVariables("HTTP_X_WAP_PROFILE")) Then
		Wap_CheckMobile="wap"
		Exit Function
	End If
	If Not IsEmpty(Request.ServerVariables("HTTP_PROFILE")) Then
		Wap_CheckMobile="wap"
		Exit Function
	End If

End Function


'初始化配置
Function WAPConfig_Initialize()
	Dim c
	Set c = New TConfig
	c.Load("Wap")
	If c.Exists("wap_version")=False Then
		c.Write "wap_version","1.0"
		c.Write "WAP_DISPLAY_COUNT","5"
		c.Write "WAP_COMMENT_COUNT","5 "
		c.Write "WAP_PAGEBAR_COUNT","5"
		c.Write "WAP_SINGLE_SIZE","2000"
		c.Write "WAP_MUTUALITY_LIMIT","5"
		c.Write "WAP_COMMENT_ENABLE","True"
		c.Write "WAP_DISPLAY_MODE_ALL","True"
		c.Write "WAP_DISPLAY_CATE_ALL","True"
		c.Write "WAP_DISPLAY_PAGEBAR_ALL","True"
		c.Write "WAP_DISABLE","True"		
		c.Save
		Call SetBlogHint_Custom("第一次安装WAP插件，已经为您导入初始配置。")
	End If
	Set c=Nothing
End Function

' Save Config to DB
Function SaveWAPConfig2DB()

	Dim c
	Set c = New TConfig
	c.Load "Wap"
	Dim i
	For Each i In Request.Form
		If Not IsEmpty(Request.Form(i)) Then
			c.Write i,Request.Form(i)
		End If
	Next
	c.Save
	Call SaveWAPConfig2Option
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "main.asp"	

End Function


' Save Config to option.asp
Function SaveWAPConfig2Option()

	Dim strContent
	strContent=LoadFromFile(BlogPath & "zb_users\plugin\wap\option_init.html","utf-8")

	Set ConfigMetas=New TMeta
	IsRunConfigs=False
	Call GetConfigs()
	
	Dim c
	Set c = New TConfig
	c.Load("Wap")
	Dim i
	For i=1 To c.Count
		If Trim(c.Meta.GetValue(c.Meta.Names(i)))="" And InStr(strContent,""""& "<#"&c.Meta.Names(i)&"#>" &"""")=0 Then
			strContent=Replace(strContent,"<#"&c.Meta.Names(i)&"#>","Empty")
		Else
			strContent=Replace(strContent,"<#"&c.Meta.Names(i)&"#>",Replace(c.Meta.GetValue(c.Meta.Names(i)),"""",""""""))
		End If
	Next

	'Call c.Save()
	Set c=Nothing
	Call SaveToFile(BlogPath & "zb_users\plugin\wap\option.asp",strContent,"utf-8",False)

End Function
'*********************************************************
%>