<!-- #include file="include_plugin.asp"-->
<!-- #include file="function.asp"-->

<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.9 其它版本的Z-blog未知
'// 插件制作:    ZSXSOFT(http://www.zsxsoft.com/)
'// 备    注:    FileManage - 挂口函数页
'///////////////////////////////////////////////////////////////////////////////

'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************


'注册插件
Call RegisterPlugin("FileManage","ActivePlugin_FileManage")
'挂口部分
Function ActivePlugin_FileManage()

	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(1,ZC_MSG210,GetCurrentHost&"zb_users/plugin/filemanage/main.asp","nav_file","aSiteFileMng",BlogHost&"zb_users/plugin/filemanage/images/folder_1.png"))
 
	Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu",MakeSubMenu("修改当前主题模板","../../ZB_USERS/plugin/FileManage/main.asp?act=ThemeEditor","m-left",False))

	
End Function
'*********************************************************
' 得到插件名
'*********************************************************
Function FileManage_GetPluginName(p,f)
	On Error Resume Next
	dim z,k,l
	z=LCase(f)
	k=LCase(p)
	l=lcase(blogpath)
	k=IIf(Right(k,1)="\",Left(k,Len(k)-1),k)
	l=IIf(Right(l,1)="\",Left(l,Len(l)-1),l)
	if k=l & "\zb_users\plugin" Then
		Select Case f
		Case "p_config.asp" FileManage_GetPluginName="总配置文件"
		Case "p_include.asp" FileManage_GetPluginName="插件include文件"
		Case "p_theme.asp" FileManage_GetPluginName="主题插件include文件"
		
		Case Else
			Dim strXmlFile,objXmlFile
			Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
			objXmlFile.async = False
			objXmlFile.ValidateOnParse=False
			strXmlFile =BlogPath & "/ZB_USERS/PLUGIN/" & f & "/" & "Plugin.xml"
			objXmlFile.load(strXmlFile)
			If objXmlFile.readyState=4 Then
				If objXmlFile.parseError.errorCode <> 0 Then
				Else
					FileManage_GetPluginName=TransferHTML(objXmlFile.documentElement.selectSingleNode("name").text,"[html-format]")
				End If
			End If
		End Select
	End If
	Set objXmlFile=Nothing
End Function
'*********************************************************
' 得到主题名
'*********************************************************
Function FileManage_GetThemeName(p,f)
	On Error Resume nEXT
	dim z,k,l
	z=LCase(f)
	k=LCase(p)
	l=lcase(blogpath)
	k=IIf(Right(k,1)="\",Left(k,Len(k)-1),k)
	l=IIf(Right(l,1)="\",Left(l,Len(l)-1),l)
	if k=l & "\zb_users\theme" Then
		Dim strXmlFile,objXmlFile
		Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		strXmlFile =BlogPath & "/ZB_USERS/THEME/" & f & "/" & "Theme.xml"
		objXmlFile.load(strXmlFile)
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode <> 0 Then
			Else
				FileManage_GetThemeName=TransferHTML(objXmlFile.documentElement.selectSingleNode("name").text,"[html-format]")
			End If
		End If
	End If
	Set objXmlFile=Nothing
End Function


%>
