<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("AdminColor","ActivePlugin_AdminColor")
'挂口部分
Function ActivePlugin_AdminColor()

	'插件最主要在这里挂接口。
	'Z-Blog可挂的接口有三类：Action、Filter、Response
	'建议参考Z-Wiki进行开发

	Call Add_Response_Plugin("Response_Plugin_Admin_SiteInfo","<p>后台配色:<hr/></p>")


	Call Add_Response_Plugin("Response_Plugin_Admin_Header","<link rel=""stylesheet"" type=""text/css"" href="""+BlogHost+"zb_users/plugin/admincolor/css.asp""/>")
	
End Function




Dim BlodColor(5)
Dim NormalColor(5)
Dim LightColor(5)
Dim HighColor(5)
Dim AntiColor(5)


BlodColor(0)="#1d4c7d"
NormalColor(0)="#3a6ea5"
LightColor(0)="#b0cdee"
HighColor(0)="#3399cc"
AntiColor(0)="#d60000"


BlodColor(1)="#143c1f"
NormalColor(1)="#5b992e"
LightColor(1)="#bee3a3"
HighColor(1)="#6ac726"
AntiColor(1)="#d60000"


%>