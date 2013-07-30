<%

Function Server_Initialize()

	If Request.QueryString("action")="update" Then intHighlight=3
	strWrite=""
	bolFrame=True
	bolPost=IIf(Request.ServerVariables("REQUEST_METHOD")="POST",True,False)
	
	Set objXmlHttp=Server.CreateObject("MSXML2.ServerXMLHTTP")
	
	Select Case Request.QueryString("action")
		Case "view"
			strURL="view.asp?"
			intHighlight=-1
		Case "catalog"
			strURL="catalog.asp?"
			If Request.QueryString("cate")=2 Then intHighlight=2
			If Request.QueryString("cate")=1 Then intHighlight=3
		Case "app"
			strURL="app.asp?"
		Case "vaildcode"
			Response.ContentType="image/gif"
			strURL="zb_system/function/c_validcode.asp?"
			bolIsBinary=True
			bolFrame=False
		Case "cmd"
			strURL="zb_system/cmd.asp?"
			bolFrame=False	
		Case "install"
			Response.Redirect "app_download.asp?url=" & Server.URLEncode(Request.QueryString("path"))
		Case "update"
			If Request.QueryString("silent")="true" Then
				If disable_check="True" Then
					Response.Write "void"
					Response.End
				End If
			End If
	
			intHighlight=4
			Call ReCheck
			strList=CheckXML()
			appcentre_updatelist=strList
			If Replace(strList,",","")<>"" Then
				strURL="app.asp?act=checkupdate&updatelist="&Server.URLEncode(strList)&"&"
			Else
				strURL="?"
			End If
			
			If Request.QueryString("silent")="true" Then 
				If CLng(appcentre_blog_last)> BlogVersion Then
					Response.Write "$('.divHeader').before('<div class=""hint""><p class=""hint hint_teal""><font color=""orangered"">Z-Blog有新版本!请立刻升级!!! <a href="""&bloghost&"zb_users/PLUGIN/AppCentre/update.asp"">升级</a></font></p></div>');"
				End If
				If Replace(appcentre_updatelist,",","")<>"" Then
					Response.Write "$('.divHeader').before('<div class=""hint""><p class=""hint hint_teal""><font color=""orangered"">发现"& UBound(Split(appcentre_updatelist,",")) &"个应用更新! <a href="""&bloghost&"zb_users/plugin/appcentre/server.asp?action=update"">更新</a></font></p></div>');"
				End If
				Response.End
			End If
	
			If Replace(strList,",","")="" Then
				Call SetBlogHint_Custom("您没有可以更新的应用.")
				Response.Redirect "server.asp"
			End If
			
		Case Else
			strURL="?"
	End Select
	
End Function 


Sub Server_SendRequest

	'On Error Resume Next
	Randomize
	strURL=APPCENTRE_URL & strURL
	strURL=strURL & Request.QueryString & "&rnd="&Rnd
	objXmlHttp.Open Request.ServerVariables("REQUEST_METHOD"),strURL
	If bolPost Then objXmlhttp.SetRequestHeader "Content-Type","application/x-www-form-urlencoded"
	objXmlhttp.SetRequestHeader "User-Agent","AppCentre/"&app_version & " ZBlog/"&BlogVersion&" "&Request.ServerVariables("HTTP_USER_AGENT") &""
	objXmlhttp.SetRequestHeader "Cookie","username="&vbsescape(login_un)&"; password="&vbsescape(login_pw)
	'为一些有趣的活动的防作弊
	objXmlhttp.SetRequestHeader "Website",ZC_BLOG_HOST
	objXmlhttp.SetRequestHeader "AppCentre",app_version
	objXmlhttp.SetRequestHeader "ZBlog",BlogVersion
	objXmlhttp.SetRequestHeader "ClientIP",GetReallyIP()
	
	objXmlHttp.Send Request.Form.Item
	
End Sub


Function AddHtml(html,stat)
	Select Case stat
	Case 0
		strResponse=Replace(strResponse,"</head>",html&"</head>")
	Case 1
		strResponse=Replace(strResponse,"</body>",html&"</body>")
	Case 2
		strResponse=Replace(strResponse,"<head>","<head>"&html)
	Case 3
		strResponse=Replace(strResponse,"<body>","<body>"&html)
	End Select
End Function


'for 2.0 users
Function GetReallyIP()

	Dim strIP
	strIP=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If strIP="" Or InStr(strIP,"unknown") Then
		strIP=Request.ServerVariables("REMOTE_ADDR")
	ElseIf InStr(strIP,",") Then
		strIP=Split(strIP,",")(0)
	ElseIf InStr(strIP,";") Then
		strIP=Split(strIP,";")(0)
	End If
	
	GetReallyIP=Trim(strIP)

End Function


Sub Server_FormatResponse()
	If objXmlHttp.ReadyState=4 Then
		If objXmlhttp.Status=200 Then
			If bolIsBinary=False Then
				strResponse=objXmlhttp.ResponseText
				strResponse=Replace(strResponse,"$bloghost$",BlogHost)
				strResponse=Replace(strResponse,"$pluginlist$",ZC_USING_PLUGIN_LIST)
				strResponse=Replace(strResponse,"$zbversion$",BlogVersion)
				strResponse=Replace(strResponse,"$appcentre$",app_version)
				strResponse=Replace(strResponse,"$username_$",login_un)
				strResponse=Replace(strResponse,"catalog.asp?","server.asp?action=catalog&")
				strResponse=Replace(strResponse,APPCENTRE_URL&"app.asp?","server.asp?action=app&")
				strResponse=Replace(strResponse,APPCENTRE_URL&"app.asp","server.asp?action=app&")
				strResponse=Replace(strResponse,APPCENTRE_URL&"view.asp?","server.asp?action=view&")
				strResponse=Replace(strResponse,APPCENTRE_URL&"""","server.asp""")
				strResponse=Replace(strResponse,APPCENTRE_URL&"zb_system/function/c_validcode.asp?name=commentvalid","server.asp?action=vaildcode")
				strResponse=Replace(strResponse,APPCENTRE_URL&"zb_system/cmd.asp?","server.asp?action=cmd&")
				Dim objRegExp
				Set objRegExp=New RegExp
				objRegExp.Pattern="<!--client_begin([\d\D]+?)-->"
				objRegExp.Global=True
				strResponse=objRegExp.Replace(strResponse,"$1")
				objRegExp.Pattern="<!--server_begin-->([\d\D]+?)<!--server_end-->"
				strResponse=objRegExp.Replace(strResponse,"")
			Else
				Response.BinaryWrite objXmlHttp.ResponseBody
				Response.End
			End If
		Else
			strResponse=ShowErr(True,"") 
		End If
	Else
		strResponse=ShowErr(True,"") 
	End If
	If Err.Number<>0 Then strResponse=ShowErr(True,"") 
End Sub

Function ShowErr(isHttp,str)
	If isHttp Then
		strWrite="<p>处理<a href="""&strURL&""" target=""_blank"">"&strURL&"</a>"
		strWrite=strWrite&"(method:"&TransferHTML(Request.ServerVariables("REQUEST_METHOD"),"[html-format]")&")时出错：</p>"
        strWrite=strWrite&"<p>ASP错误信息：" & IIf(Err.Number=0, "无" , Err.Number & "(" & Err.Description & ")" ) & "</p>"
        strWrite=strWrite&"<p>HTTP状态码："
		If objXmlhttp.ReadyState<4 Then
			strWrite=strWrite & "未发送请求"
		Else
			strWrite=strWrite & objXmlhttp.Status
		End If
		strWrite=strWrite&"</p>"
        strWrite=strWrite&"<p>&nbsp;</p>"
        strWrite=strWrite&"<p>可能的原因有：</p>"
        strWrite=strWrite&"    <p>"
        strWrite=strWrite&"    <ol>"
        strWrite=strWrite&"      <li>您的服务器不允许通过HTTP协议连接到：<a href="""&APPCENTRE_URL&""" target=""_blank"">"&APPCENTRE_URL&"</a>；</li>"
        strWrite=strWrite&"      <li>您进行了一个错误的请求；</li>"
        strWrite=strWrite&"      <li>服务器暂时无法连接，可能是遭到攻击或者检修中。</li>"
        strWrite=strWrite&"    </ol>"
        strWrite=strWrite&"    <p>请<a href=""javascript:location.reload()"">点击这里刷新重试</a>，或者到<a href=""http://bbs.rainbowsoft.org"" target=""_blank"">Z-Blogger论坛</a>发帖询问。</p>"
    Else
        strWrite=str
    End If
	ShowErr=strWrite
End Function



%>