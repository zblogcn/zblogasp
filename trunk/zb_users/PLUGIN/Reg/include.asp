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
Call RegisterPlugin("Reg","ActivePlugin_Reg")
'挂口部分
Function ActivePlugin_Reg()

End Function

Function InstallPlugin_Reg()

	Call GetFunction()
	Functions(FunctionMetas.GetValue("navbar")).Content=Functions(FunctionMetas.GetValue("navbar")).Content & "<li><a href=""<#ZC_BLOG_HOST#>zb_users/plugin/reg/reg.asp"">注册</a></li>"
	Functions(FunctionMetas.GetValue("navbar")).Save

	Call ClearGlobeCache
	Call LoadGlobeCache

End Function


Function UninstallPlugin_Reg()

	Call GetFunction()
	Functions(FunctionMetas.GetValue("navbar")).Content=RemoveLibyUrl(Functions(FunctionMetas.GetValue("navbar")).Content,"<#ZC_BLOG_HOST#>zb_users/plugin/reg/reg.asp")


	Functions(FunctionMetas.GetValue("navbar")).Save

	Call ClearGlobeCache
	Call LoadGlobeCache

End Function
%>