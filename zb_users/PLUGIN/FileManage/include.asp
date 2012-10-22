<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.9 其它版本的Z-blog未知
'// 插件制作:    ZSXSOFT(http://www.zsxsoft.com/)
'// 备    注:    FileManage - 挂口函数页
'///////////////////////////////////////////////////////////////////////////////

'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************
Const FileManage_ShowPluginName=True
Const FileManage_ShowThemesName=True
Const FileManage_CodeMirror=False

'注册插件
Call RegisterPlugin("FileManage","ActivePlugin_FileManage")
'挂口部分
Function ActivePlugin_FileManage()

	Call Add_Response_Plugin("Response_Plugin_Admin_Left",MakeLeftMenu(1,ZC_MSG210,GetCurrentHost&"zb_users/plugin/filemanage/main.asp","nav_file","aSiteFileMng",""))
 
	Call Add_Response_Plugin("Response_Plugin_ThemeMng_SubMenu",MakeSubMenu("修改当前主题模板","../../ZB_USERS/plugin/FileManage/main.asp?act=ThemeEditor","m-left",False))
	If FileManage_ShowPluginName=True Then Call Add_Action_Plugin("Action_Plugin_FileManage_ExportInformation_NotFound","FileManage_GetPluginName(""{path}"",""{f}"")")
	If FileManage_ShowThemesName=True Then Call Add_Action_Plugin("Action_Plugin_FileManage_ExportInformation_NotFound","FileManage_GetThemeName(""{path}"",""{f}"")")
	
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
	if k=l & "\zb_users\plugin" Then
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