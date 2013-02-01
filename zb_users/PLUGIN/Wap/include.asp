<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("Wap","ActivePlugin_Wap")
'挂口部分
Function ActivePlugin_Wap()

End Function

Call Add_Action_Plugin("Action_Plugin_Default_WithOutConnect_Begin","If ZC_DISPLAY_COUNT_WAP>0 Then If CheckMobile() Then Server.Transfer ""zb_users\plugin\wap\wap.asp"":Response.End")


'*********************************************************
' 目的：   检查是否手机端访问
'*********************************************************
Function CheckMobile()
	
	
	Dim bolCheck
	bolCheck=True
	'是否由wap转入电脑版
	If Request.Cookies("CheckMobile")="false" Then
		CheckMobile=False
		bolCheck=False
	End If
	If  Not IsEmpty(Request.ServerVariables("HTTP_REFERER"))  And  InStr(LCase(Request.ServerVariables("HTTP_REFERER")),ZC_FILENAME_WAP) Then 
			CheckMobile=False
			Response.Cookies("CheckMobile")="false"
			bolCheck=False
	End If 
	'是否专用wap浏览器
	If InStr(LCase(Request.ServerVariables("HTTP_ACCEPT")), "application/vnd.wap.xhtml+xml") Or Not IsEmpty(Request.ServerVariables("HTTP_X_PROFILE")) Or Not IsEmpty(Request.ServerVariables("HTTP_PROFILE")) Then
			CheckMobile=True
			ZC_ISWAP=True
			Exit Function
	End If 
	
		'是否（智能）手机浏览器
	Dim MobileBrowser_List,PCBrowser_List,UserAgent
	MobileBrowser_List ="up.browser|up.link|mmp|iphone|android|wap|netfront|java|opera\smini|ucweb|windows\sce|symbian|series|webos|sonyericsson|sony|blackberry|cellphone|dopod|nokia|samsung|palmsource|palmos|pda|xphone|xda|smartphone|pieplus|meizu|midp|cldc|brew|tear|ipad|kindle"
	PCBrowser_List="mozilla|chrome|safari|opera|m3gate|winwap|openwave"
	UserAgent = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
	If CheckRegExp(UserAgent,MobileBrowser_List) Then 
		If bolCheck Then CheckMobile=True
		ZC_ISWAP=True
		Exit Function
	ElseIf CheckRegExp(UserAgent,PCBrowser_List) Then '未知手机浏览器，其UA标识为常见浏览器，不跳转
		If bolCheck Then CheckMobile=False
		ZC_ISWAP=False
		Exit Function
	Else 
		ZC_ISWAP=False
		If bolCheck Then CheckMobile=False 
	End If 

End Function 
'*********************************************************

%>