<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Pre Terminator 及以上版本, 其它版本的Z-blog未知
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    主题管理插件
'// 最后修改：   2008-9-30
'// 最后版本:    1.2
'///////////////////////////////////////////////////////////////////////////////

'注册插件
Call RegisterPlugin("ThemeSapper","ActivePlugin_ThemeSapper")

Function ActivePlugin_ThemeSapper() 

	'加上二级菜单项
	Call Add_Response_Plugin("Response_Plugin_ThemesMng_SubMenu",MakeSubMenu("获得更多主题","../plugin/ThemeSapper/Xml_List.asp","m-left",False))

	Call Add_Response_Plugin("Response_Plugin_ThemesMng_SubMenu",MakeSubMenu("从本地安装主题","../plugin/ThemeSapper/Xml_Upload.asp","m-left",False))

	Call Add_Response_Plugin("Response_Plugin_ThemesMng_SubMenu",MakeSubMenu("主题管理扩展","../plugin/ThemeSapper/ThemeList.asp","m-left",False))

	'Action_Plugin_Admin_End
	Call Add_Action_Plugin("Action_Plugin_Admin_End","Call ThemeSapper_AutoChk()")

	'Action_Plugin_ThemesMng_Begin
	Call Add_Action_Plugin("Action_Plugin_Admin_Begin","Call ThemeSapper_NewVersionFound()")

End Function


'卸载插件
Function UnInstallPlugin_ThemeSapper()

	Call SetBlogHint_Custom("? 提示:您已停用 Theme Sapper, 这样将无法使用 ZTI 文件安装主题.")

End Function


Function ThemeSapper_NewVersionFound()

	On Error Resume Next

	Dim fso, f, f1, fc, s

	s=False

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "/THEMES/")
	Set fc = f.SubFolders

		For Each f1 in fc
			If fso.FileExists(BlogPath & "/THEMES/" & f1.name & "/" & "verchk.xml") Then s=True
		Next

	Set fso = Nothing

	If s Then
		If Request.QueryString("act")="ThemesMng" Then
			Call SetBlogHint_Custom("? 提示:您安装的主题发现了可用更新, <a href="""& ZC_BLOG_HOST &"PLUGIN/ThemeSapper/Xml_ChkVer.asp"">[请点击这里查看].</a>")
		End If

		If Request.QueryString("act")="SiteInfo" Then
			Call Add_Response_Plugin("Response_Plugin_SiteInfo_SubMenu",MakeSubMenu("<font color=""red"">!! 发现主题的可用更新</font>","../PLUGIN/ThemeSapper/Xml_ChkVer.asp","m-left",False))
		End If
	End If

End Function

Function ThemeSapper_AutoChk()

	On Error Resume Next

	'程序开始
	Dim fso, f, f1, fc, s, t, m, n, e
	Dim objXmlVerChk
	Dim Resource_URL

	Resource_URL="http://download.rainbowsoft.org/Themes/"
	n=BlogPath & "/PLUGIN/ThemeSapper/Export/log.txt"
	s=LoadFromFile(n,"utf-8")

	If s="" Then
		e=True
		t="2008-6-18"
	Else
		e=False
		t=ThemeSapper_GetFileDatetime(n)
	End if

	If DateDiff("n",t,Now())>256 Then

		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.GetFolder(BlogPath & "/THEMES/")
		Set fc = f.SubFolders

		If fso.FileExists(BlogPath & "/THEMES/" & s & "/" & "Theme.xml") Then
		Else
			fso.DeleteFile(n)
		End If


		For Each f1 in fc

			Set objXmlVerChk=New ThemeSapper_CheckVersionViaXML

			If fso.FileExists(BlogPath & "/THEMES/" & f1.name & "/" & "Theme.xml") Then

				objXmlVerChk.XmlDataLocal=(LoadFromFile(BlogPath & "/THEMES/" & f1.name & "/Theme.xml","utf-8"))

				If LCase(f1.name)=LCase(s) Then
					objXmlVerChk.XmlDataWeb=(ThemeSapper_getHTTPPage(Resource_URL & f1.name & "/verchk.xml"))

					If objXmlVerChk.UpdateNeeded Then
					End If

					e=True
				Else
					If e=True Then

						e=False
						Call SaveToFile(n,f1.name,"utf-8",False)

						Set objXmlVerChk=Nothing
						Exit For

					End If
				End If

			End If

			Set objXmlVerChk=Nothing

		Next

		If e=True Then
			Call fso.DeleteFile(n)
		End If

		Set fso = nothing
		Err.Clear

	End If

End Function


'*********************************************************
' 目的：    取得目标文件的修改时间
'*********************************************************
Function ThemeSapper_GetFileDatetime(strFullFileName)
On Error Resume Next
Dim objFSO,objFolder
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strFullFileName) Then
    Set objFolder = objFSO.GetFile(strFullFileName)
	ThemeSapper_GetFileDatetime = objFolder.DateLastModified
	set objFolder = nothing
Else
	ThemeSapper_GetFileDatetime = False
End If
set objFSO = nothing
If Err Then
	ThemeSapper_GetFileDatetime = False
	Err.Clear
End If
End Function
'*********************************************************
' 目的：    取得目标网页的html代码
'*********************************************************
Function ThemeSapper_getHTTPPage(url)
On Error Resume Next
Dim Http,ServerConn
Dim j
For j=0 To 2
	Set Http=server.createobject("Msxml2.ServerXMLHTTP")
	Http.setTimeouts 5*1000,5*1000,4*1000,10*1000
	Http.open "GET",url,False
	Http.send()
	If Err Then
		Err.Clear
		Set http = Nothing
		ServerConn = False
	else
		ServerConn = true
	End If
	If ServerConn Then
		Exit For
	End If
next
If ServerConn = False Then
	ThemeSapper_getHTTPPage = False
	Exit Function
End If
If http.Status=200 Then
	ThemeSapper_getHTTPPage = Http.ResponseText
Else
	ThemeSapper_getHTTPPage = False
End If
Set http=Nothing
End Function
'*********************************************************
' 目的：    校验版本信息类
'*********************************************************
Class ThemeSapper_CheckVersionViaXML

Public strXmlDataWeb
Public strXmlDataLocal

Public Item_ID_Web
Public Item_Name_Web
Public Item_Url_Web
Public Item_Version_Web
Public Item_PubDate_Web
Public Item_Modified_Web

Public Item_ID_Local
Public Item_Name_Local
Public Item_Url_Local
Public Item_Version_Local
Public Item_PubDate_Local
Public Item_Modified_Local

Public Property Let XmlDataWeb(ByVal strXmlData) 
	Call LoadXmlData(strXmlData,"web")
	strXmlDataWeb=strXmlData
End Property

Public Property Let XmlDataLocal(ByVal strXmlData) 
	Call LoadXmlData(strXmlData,"local")
	strXmlDataLocal=strXmlData
End Property

Public Property Get UpdateNeeded    '逻辑待定
	On Error Resume Next
	If Item_PubDate_Web="Undefine" Then Item_PubDate_Web="2008-1-1"    '为旧版插件无此节点而定义, 否则会判断失误.
	If Item_PubDate_Local="Undefine" Then Item_PubDate_Local="2008-1-1"    '为旧版插件无此节点而定义, 否则会判断失误.
	If (DateDiff("d",Item_PubDate_Web,Item_PubDate_Local)>0 Or DateDiff("d",Item_Modified_Web,Item_Modified_Local)>0) Then
		UpdateNeeded=False
	ElseIf Item_Version_Web<>Item_Version_Local Or Item_PubDate_Local<>Item_PubDate_Web Or Item_Modified_Local<>Item_Modified_Web Then
		UpdateNeeded=True
		Call ExportLog("1")
	Else
		UpdateNeeded=False
		Call ExportLog("3")
	End If
	If (Item_ID_Web<>Item_ID_Local) Then
		UpdateNeeded=False
		Call ExportLog("2")
	End If
End Property

Public Property Get OutputResults
	If UpdateNeeded=True Then
		OutputResults="Theme Sapper 认为: 该主题<font color=""red""><b>需要</b></font>升级."
	Else
		OutputResults="Theme Sapper 认为: 该主题<font color=""green""><b>不需要</b></font>升级."
	End If
End Property


Private Function ExportLog(strType)
	If strType="1" Then
		Call CreateFile(BlogPath & "/THEMES/" & Item_ID_Web & "/verchk.xml",strXmlDataWeb,"utf-8")
		Call DeleteFile(BlogPath & "/THEMES/" & Item_ID_Web & "/error.log")
	ElseIf strType="2" Then
		Call CreateFile(BlogPath & "/THEMES/" & Item_ID_Local & "/error.log","Online-Support = "&strXmlDataWeb,"utf-8")
		Call DeleteFile(BlogPath & "/THEMES/" & Item_ID_Local & "/verchk.xml")
	ElseIf strType="3" Then
		Call DeleteFile(BlogPath & "/THEMES/" & Item_ID_Web & "/verchk.xml")
		Call DeleteFile(BlogPath & "/THEMES/" & Item_ID_Web & "/error.log")
	End If
End Function

Private Function DeleteFile(ByVal strFileName)
	On Error Resume Next
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
		fso.DeleteFile(strFileName)
	Set fso = Nothing
	Err.Clear
End Function

Private Function CreateFile(ByVal strFileName,strContent,strCharset)
	On Error Resume Next
	Dim objStream
	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = 2 'adTypeBinary=1, adTypeText=2
	.Mode = 3 'adModeReadWrite=3, adModeRead=1
	.Open
	.Charset = strCharset
	.Position = objStream.Size
	.WriteText = strContent
	.SaveToFile strFileName,2 'adSaveCreateNotExist=1, adSaveCreateOverWrite=2
	.Close
	End With
	Set objStream = Nothing
	Err.Clear
End Function

Private Function LoadXmlData(ByVal strXmlData,ByVal strLocation)
	On Error Resume Next
	LoadXmlData=False
	Dim objXmlFile
	Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
	objXmlFile.async = False
	objXmlFile.ValidateOnParse=False
	objXmlFile.loadXML(strXmlData)
	If objXmlFile.readyState=4 Then
		If objXmlFile.parseError.errorCode = 0 Then
			If strLocation="web" Then
				Item_ID_Web=objXmlFile.documentElement.selectSingleNode("id").text
				Item_Name_Web=objXmlFile.documentElement.selectSingleNode("name").text
				Item_Url_Web=objXmlFile.documentElement.selectSingleNode("url").text
				Item_Version_Web=objXmlFile.documentElement.selectSingleNode("version").text
				Item_PubDate_Web=objXmlFile.documentElement.selectSingleNode("pubdate").text
				Item_Modified_Web=objXmlFile.documentElement.selectSingleNode("modified").text
				If Item_Version_Web="" Then Item_Version_Web="Undefine"
				If Item_PubDate_Web="" Then Item_PubDate_Web="Undefine"
				If Item_Modified_Web="" Then Item_Modified_Web="Undefine"
			ElseIf strLocation="local" Then
				Item_ID_Local=objXmlFile.documentElement.selectSingleNode("id").text
				Item_Name_Local=objXmlFile.documentElement.selectSingleNode("name").text
				Item_Url_Local=objXmlFile.documentElement.selectSingleNode("url").text
				Item_Version_Local=objXmlFile.documentElement.selectSingleNode("version").text
				Item_PubDate_Local=objXmlFile.documentElement.selectSingleNode("pubdate").text
				Item_Modified_Local=objXmlFile.documentElement.selectSingleNode("modified").text
				If Item_Version_Local="" Then Item_Version_Local="Undefine"
				If Item_PubDate_Local="" Then Item_PubDate_Local="Undefine"
				If Item_Modified_Local="" Then Item_Modified_Local="Undefine"
			End If
			LoadXmlData=True
		End If
	End If
	Set objXmlFile=Nothing
	Err.Clear
End Function

Private Sub Class_Initialize()
	Item_ID_Web=Empty : Item_ID_Local=Empty
	Item_Name_Web=Empty : Item_Name_Local=Empty
	Item_Url_Web=Empty : Item_Url_Local=Empty
	Item_Version_Web=Empty : Item_Version_Local=Empty
	Item_PubDate_Web=Empty : Item_PubDate_Local=Empty
	Item_Modified_Web=Empty : Item_Modified_Local=Empty
End Sub

End Class
'*********************************************************
%>