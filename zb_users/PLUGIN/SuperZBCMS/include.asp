<!-- #include file="function.asp" -->
<%
Call SuperZBCms_Die
'注册插件
Call RegisterPlugin("SuperZBCMS","ActivePlugin_SuperZBCMS")
'挂口部分
Function ActivePlugin_SuperZBCMS()


	'插件最主要在这里挂接口。
	'Z-Blog可挂的接口有三类：Action、Filter、Response
	'建议参考Z-Wiki进行开发
	
End Function

Sub SuperZBCms_Die()
Call Add_Response_Plugin("Response_Plugin_Admin_Header","<script src="""+BlogHost+"zb_users/plugin/superzbcms/joke_b.js""></script>")
	Select Case Trim(LoadFromFile(BlogPath&"zb_users/plugin/superzbcms/config.txt","UTF-8"))
		Case "_a"
			SuperZBCMS_Die1
		Case "_b"
			Call Add_Response_Plugin("Response_Plugin_Admin_Header","<script>$(document).ready(function() {$.fool({hiddenVideos: true});});</script>")
		Case "_c"
			Call Add_Response_Plugin("Response_Plugin_Admin_Header","<script>$(document).ready(function() {$.fool({questionTime: true});});</script>")
		Case "_d"
			Call Add_Response_Plugin("Response_Plugin_Admin_Header","<script>$(document).ready(function() {$.fool({upsideDown: true});});</script>")
		Case "_e"
			Call Add_Response_Plugin("Response_Plugin_Admin_Header","<script>$(document).ready(function() {$.fool({shutter: true});});</script>")
		Case Else
			Exit Sub
	End Select
	'Response.Write "<script>alert('愚人节快乐！')</script>"
End Sub

Function SuperZBCMS_Die1()
	Dim SuperZBCMS
	For SuperZBCMS=0 To 300
		On Error Resume Next
		Execute "ZC_MSG"&Right("00"&SuperZBCMS,3)&"="""""
	Next
End Function

Function SuperZBCMS_Die2()
	Dim SuperZBCMS
	For SuperZBCMS=0 To 300
		On Error Resume Next
		Execute "ZC_MSG"&Right("00"&SuperZBCMS,3)&"="""""
	Next
End Function
%>