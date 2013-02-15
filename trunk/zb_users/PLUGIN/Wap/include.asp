﻿<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("Wap","ActivePlugin_Wap")
'挂口部分
Function ActivePlugin_Wap()



End Function

Call Add_Action_Plugin("Action_Plugin_Default_WithOutConnect_Begin","Wap_Check()")

Dim Wap_Type


Dim WAP_DISPLAY_COUNT
WAP_DISPLAY_COUNT=5

Dim WAP_COMMENT_COUNT
WAP_COMMENT_COUNT=5

Dim WAP_PAGEBAR_COUNT
WAP_PAGEBAR_COUNT=5

Dim WAP_SINGLE_SIZE
WAP_SINGLE_SIZE=1000

Dim WAP_SINGLE_PAGEBAR_COUNT
WAP_SINGLE_PAGEBAR_COUNT=5

Dim WAP_FILENAME
WAP_FILENAME="wap.asp"

Dim WAP_COMMENT_ENABLE
WAP_COMMENT_ENABLE=True

'全文
Dim WAP_DISPLAY_MODE_ALL
WAP_DISPLAY_MODE_ALL=True

'显示分类导航
Dim WAP_DISPLAY_CATE_ALL
WAP_DISPLAY_CATE_ALL=True

'分页条
Dim WAP_DISPLAY_PAGEBAR_ALL
WAP_DISPLAY_PAGEBAR_ALL=True

'相关文章数量
Dim WAP_MUTUALITY_LIMIT
WAP_MUTUALITY_LIMIT=5






Function Wap_Check()

	Dim s

	If Request.QueryString("mod")="pc" Then
		Exit Function
	End If

	If Request.QueryString("mod")="pad" Then
		Server.Transfer "zb_users\plugin\wap\pad.asp":Response.End
	End If

	If Request.QueryString("mod")="wap" Then
		Server.Transfer "zb_users\plugin\wap\wap.asp":Response.End
	End If

	Wap_Type=Wap_CheckMobile()

	If Wap_Type="pad" Then Server.Transfer "zb_users\plugin\wap\pad.asp":Response.End

	If Wap_Type="wap" Then Server.Transfer "zb_users\plugin\wap\wap.asp":Response.End

End Function 


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
%>