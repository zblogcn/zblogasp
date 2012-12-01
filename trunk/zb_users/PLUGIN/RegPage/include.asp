<!-- #include file="include_plugin.asp"-->
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.9 其它版本的Z-blog未知
'// 插件制作:    ZSXSOFT(http://www.zsxsoft.com/)
'// 备    注:    Reg - 挂口函数页
'///////////////////////////////////////////////////////////////////////////////

'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************


'注册插件
Call RegisterPlugin("RegPage","ActivePlugin_RegPage")
'挂口部分
Function ActivePlugin_RegPage()
	
End Function

Function InstallPlugin_RegPage()

	Call GetFunction()
	Functions(FunctionMetas.GetValue("navbar")).Content=Functions(FunctionMetas.GetValue("navbar")).Content & "<li><a href=""<#ZC_BLOG_HOST#>zb_users/plugin/regpage/reg.asp"">注册</a></li>"
	Functions(FunctionMetas.GetValue("navbar")).Save

	Call ClearGlobeCache
	Call LoadGlobeCache

	Dim a
	Set a=New TConfig
	a.Load "RegPage"
	If a.Exists("Version")=False Then
		a.Write "Level",4
		a.Write "Version","1.0"
		a.Save
	End If
End Function


Function UninstallPlugin_RegPage()

	Call GetFunction()
	Functions(FunctionMetas.GetValue("navbar")).Content=RemoveLibyUrl(Functions(FunctionMetas.GetValue("navbar")).Content,"<#ZC_BLOG_HOST#>zb_users/plugin/regpage/reg.asp")


	Functions(FunctionMetas.GetValue("navbar")).Save

	Call ClearGlobeCache
	Call LoadGlobeCache

End Function
%>