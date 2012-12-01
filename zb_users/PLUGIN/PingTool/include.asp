<%

'注册插件
Call RegisterPlugin("pingtool","ActivePlugin_pingtool")
Dim PingTool_PingEnable
'具体的接口挂接
Function ActivePlugin_PingTool() 
	'挂上接口
	Call Add_Action_Plugin("Action_Plugin_ArticlePst_Begin","PingTool_PingEnable=Request.Form(""PingTool_PingEnable""):Call PingTool_Main()")
	Call Add_Action_Plugin("Action_Plugin_Edit_ueditor_Begin","Call PingTool_addForm()")
End Function
Sub InstallPlugin_PingTool
	Dim objConfig
	Set objConfig = New TConfig
	objConfig.Load("PingTool")
	If objConfig.Exists("Version")=False Then
		objConfig.Write "Version","1.5"
		objConfig.Save
	End If
End Sub
Function PingTool_addForm()
Call Add_Response_Plugin("Response_Plugin_Edit_Form3","<input type=""checkbox""  name=""PingTool_PingEnable"" id=""PingTool_PingEnable"" onclick="""" value=""True"" checked=""checked""/><label for=""PingTool_PingEnable"">通知Ping中心.</label><br/>")
End Function
Function PingTool_getArticle(ByRef objArticle) 
	Call Add_Action_Plugin("Action_Plugin_ArticlePst_Succeed","Call addBatch(""PingTool"",""PingTool_gotoPing("&objArticle.ID&")"")")
End Function
Function PingTool_gotoPing(a) 
	Dim objArticle
	Set objArticle=New TArticle
	objArticle.LoadInfoByID a
	If objArticle.ID>0 Then
		Dim objConfig
		Set objConfig=New TConfig
		objConfig.Load "PingTool"
		SendPing objConfig.Read("Content"),objArticle.HTMLURL
		Set objConfig=Nothing
	End If
End Function
Function PingTool_Main()
	If IsEmpty(PingTool_PingEnable)=True Then
		PingTool_PingEnable=False
	Else
		PingTool_PingEnable=True
	End If
	If PingTool_PingEnable Then
		Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Core","PingTool_getArticle")
	End If
End Function
Function SendPing_Single(url,url2)
	'On Error Resume Next
	Dim s
	s = "<?xml version=""1.0""?><methodCall><methodName>weblogUpdates.ping</methodName><params><param><value>"&TransferHTML(ZC_BLOG_NAME,"[<][>][&][""]")&"</value></param><param><value>"&url2&"</value></param></params></methodCall>"
	Response.Write "<p>发送Ping到:" & Url & "</p>"
	Response.Flush
	Dim objPing
	Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP")
	objPing.SetTimeOuts 10000, 10000, 10000, 10000 
	objPing.open "POST",url,False
	objPing.setRequestHeader "Content-Type", "text/xml"
	objPing.send s
	Set objPing = Nothing
	Err.Clear
End Function
Function SendPing(PingContent,Url2)
	Dim Url,Urls
	Urls=Split(Replace(PingContent,vbCr,""),vbLf)
	For Each Url In Urls
		If Trim(Url)<>"" Then
			Call SendPing_Single(url,url2)
		End If
	Next
End Function
%>