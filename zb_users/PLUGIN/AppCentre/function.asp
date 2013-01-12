<%
Const APPCENTRE_URL="http://app.rainbowsoft.org/"
Const APPCENTRE_UPDATE_URL="http://app.rainbowsoft.org/appcentre.asp?act=checkupdate"

Const APPCENTRE_SYSTEM_UPDATE="http://update.rainbowsoft.org/zblog2/"

Dim app_config
Dim login_un,login_pw,disableupdate_theme
Dim Pack_For,Pack_Type

Dim app_id
Dim app_name
Dim app_url
Dim app_note

Dim app_author_name
Dim app_author_email
Dim app_author_url

Dim app_source_name
Dim app_source_email
Dim app_source_url

Dim app_plugin_path
Dim app_plugin_include
Dim app_plugin_level

Dim app_adapted
Dim app_version
Dim app_pubdate
Dim app_modified
Dim app_description
Dim app_price

Dim app_dependency
Dim app_rewritefunctions
Dim app_conflict

Dim app_path

Dim app_sidebars

Dim aryDownload(),aryName()

Redim aryDownload(0)
Redim aryName(0)


Function AppCentre_Update_Restore(build,file,crc32_)

	'On Error Resume Next

	Dim objPing
	Set objPing = Server.CreateObject("Microsoft.XMLHTTP")

	objPing.open "GET", APPCENTRE_SYSTEM_UPDATE & "?" & build & "\" & file ,False
	objPing.setRequestHeader "accept-encoding", "gzip, deflate"  
	objPing.send ""

	
	Dim MyStream
    Set MyStream=Server.CreateObject("Adodb.Stream") 
	MyStream.Type = adTypeBinary
	MyStream.Mode = adModeReadWrite
    MyStream.Open 
    MyStream.Write objPing.responsebody
	
	If CRC32(MyStream)=crc32_ Then
	    MyStream.SaveToFile BlogPath & file ,2
	Else
		Response.End
	End If



	If Err.Number=0 Then
		AppCentre_Update_Restore="下载成功."
	Else
		Response.End
	End If

End Function


Function AppCentre_Update_Download(file)

	On Error Resume Next

	Dim objPing
	Set objPing = Server.CreateObject("Microsoft.XMLHTTP")

	objPing.open "GET", APPCENTRE_SYSTEM_UPDATE & "?" & file & ".xml",False
	objPing.setRequestHeader "accept-encoding", "gzip, deflate"  
	objPing.send ""

	Dim MyStream
    Set MyStream=Server.CreateObject("Adodb.Stream") 
	MyStream.Type = adTypeBinary
	MyStream.Mode = adModeReadWrite
    MyStream.Open 
    MyStream.Write objPing.responsebody
    MyStream.SaveToFile BlogPath & "zb_users/cache/update.xml" ,2


	If Err.Number=0 Then
		AppCentre_Update_Download="下载成功."
	Else
		Response.End
	End If

End Function

Function AppCentre_Update_Install()

	On Error Resume Next

	Dim objXmlFile,strXmlFile
	Dim fso, f, f1, fc, s
	Set fso = CreateObject("Scripting.FileSystemObject")

	strXmlFile =BlogPath & "zb_users/cache/update.xml"
	
	If fso.FileExists(strXmlFile) Then

		Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		objXmlFile.load(strXmlFile)
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode <> 0 Then
			Else



				Dim objXmlFiles,item,objStream
				Set objXmlFiles=objXmlFile.documentElement.SelectNodes("file")
				for each item in objXmlFiles
				Set objStream = CreateObject("ADODB.Stream")
					With objStream
					.Type = 1
					.Mode = 3
					.Open
					.Write item.nodeTypedvalue
					
					Dim i,j,k,l
					i=item.getAttributeNode("name").Value

					j=Left(i,InstrRev(i,"\"))
					k=Replace(i,j,"")
					Call CreatDirectoryByCustomDirectory("" & j)

					.SaveToFile BlogPath & "" & item.getAttributeNode("name").Value,2

					's=s& "释放 " & k & ";"
					.Close
					End With
					Set objStream = Nothing
					l=l+1
				next
				s=s& "释放 " & l & " 个文件;<br/>"
				s=s& "升级成功!!!<br/>"

				Call DelToFile(BlogPath & "zb_users/cache/update.xml")
			End If
		End If
	End If



	If Err.Number=0 Then
		AppCentre_Update_Install=s
	Else
		Response.End
	End If

End Function


Function AppCentre_CheckSystemIndex(build)

	Dim objXmlFile,strXmlFile
	Dim fso, f, f1, fc, s
	Set fso = CreateObject("Scripting.FileSystemObject")

	f=False
	
	If fso.FileExists(BlogPath & "zb_users/cache/"&build&".xml") Then

		strXmlFile =BlogPath & "zb_users/cache/"&build&".xml"

		Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		objXmlFile.load(strXmlFile)
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode <> 0 Then
			Else
				f=True
			End If
		End If
	End If

	If f=False Then

		Dim objPing
		Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP")

		objPing.open "GET", APPCENTRE_SYSTEM_UPDATE & build & ".xml",False
		objPing.send ""

		Call SaveToFile(BlogPath & "zb_users/cache/"&build&".xml",objPing.responseText,"utf-8",False)

	End If

End Function



Function AppCentre_CheckSystemLast()

	Dim objPing
	Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP")

	objPing.open "GET", APPCENTRE_SYSTEM_UPDATE,False
	objPing.send ""

	AppCentre_CheckSystemLast=objPing.responseText


	Set objPing = Nothing
End Function


Function ReCheck()
	Dim objXmlHttp,strURL,bolPost,str,bolIsBinary
	Set objXmlHttp=Server.CreateObject("MSXML2.ServerXMLHTTP")

	strUrl=APPCENTRE_UPDATE_URL&"&tname="&Server.URLEncode(Join(GetAllThemeName,","))&"&pname="&Server.URLEncode(Replace(ZC_USING_PLUGIN_LIST,"|",","))&"&rnd="&Rnd
	objXmlHttp.Open "GET",strURL
	objXmlHttp.Send 

	If objXmlHttp.ReadyState=4 Then
		If objXmlhttp.Status=200 Then
		Else
			Call ShowErr(True,"")
		End If
		
		
	Else
		ShowErr
	End If
	If Err.Number<>0 Then Call ShowErr(True,"")
	
	Call SaveToFile(BlogPath&"zb_users\cache\appcentre_plugin.xml",objXmlHttp.ResponseText,"utf-8",False)
End Function


Function CheckXml()
	Dim strTemp,strName,strType,dtmModified,dtmLocalModified
	Dim objXml,objXml2,objChildXml,objAppXml,i,objFso,j
	Set objXml=Server.CreateObject("Microsoft.XMLDOM")
	objXml.Load BlogPath&"zb_users\cache\appcentre_plugin.xml"
	If objXml.ReadyState=4 Then
		'这里该显示更新列表了
		If objXml.parseError.errorCode=0 Then
			Set objChildXml=objXml.selectNodes("//data/apps/app")
			For i=0 To objChildXml.length-1
				Set objAppXml=objChildXml(i)'
				If CLng(objAppXml.getAttributeNode("zbversion").value)<=BlogVersion Then
					strName=objAppXml.getAttributeNode("name").value
					strType=objAppXml.getAttributeNode("type").value
					dtmModified=CDate(objAppXml.getAttributeNode("modified").value)
					Set objXml2=Server.CreateObject("Microsoft.XMLDOM")
					objXml2.Load BlogPath&"zb_users\"&strType&"\"&strName&"\"&strType&".xml"
					
					If objXml2.ReadyState=4 Then
						If objXml2.parseError.errorCode=0 Then
							dtmLocalModified=CDate(objXml2.documentElement.selectSingleNode("modified").text)
						End If
					End If
					If DateDiff("d","1970-1-1 08:00",dtmModified)>DateDiff("d","1970-1-1 08:00",dtmLocalModified) Then
						Redim Preserve aryDownload(Ubound(aryDownload)+1)
						Redim Preserve aryName(Ubound(aryName)+1)
						aryDownload(Ubound(aryDownload))=objAppXml.getAttributeNode("url").value
						aryName(Ubound(aryName))=strName						
					End If
				End If
			Next
			j=objXml.SelectSingleNode("//data/blog").text
			Application.Lock
			Application("APPCENTRE_BLOG_LAST")=CStr(j)
			Application.UnLock
		End If
	End If

	CheckXml=Join(aryName,",")

End Function


Function GetAllThemeName()
	Dim aryTheme,bolFound
	bolFound=False
	aryTheme=Split(LCase(disableupdate_theme),":")
	Dim aryReturn()
	Redim aryReturn(0)
	Dim fso,f,fc,f1,f2
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(BlogPath & "zb_users/theme" & "/")
	Set fc = f.SubFolders
	For Each f1 In fc
		For f2=0 To Ubound(aryTheme)
			If aryTheme(f2)=LCase(f1.Name) Then bolFound=True:Exit For
		Next
		If bolFound=False Then
			Redim Preserve aryReturn(Ubound(aryReturn)+1)
			aryReturn(Ubound(aryReturn))=f1.Name
		End If
		bolFound=False
	Next
	GetAllThemeName=aryReturn
End Function

Sub AppCentre_SubMenu(id)
	Dim aryName,aryValue,aryPos
	aryName=Array("在线安装应用","新建插件","新建主题","开发者登录","主题列表","插件列表","检查应用更新","系统更新与文件校验")
	aryValue=Array("server.asp","plugin_edit.asp","theme_edit.asp","login.asp","server.asp?action=catalog&cate=2","server.asp?action=catalog&cate=1","server.asp?action=update","update.asp")
	aryPos=Array("m-left","m-right","m-right","m-right","m-left","m-left","m-left","m-left")
	Dim i 
	For i=0 To Ubound(aryName)
		Response.Write MakeSubMenu(aryName(i),aryValue(i),aryPos(i) & IIf(id=i," m-now",""),False)
	Next
End Sub

Sub AppCentre_InitConfig
	Set app_config=New TConfig
	app_config.Load "AppCentre"
	login_un=app_config.read("DevelopUserName")
	login_pw=app_config.read("DevelopPassWord")
	disableupdate_theme=app_config.read("DisableUpdateTheme")
End Sub

Function AppCentre_GetLastModifiTime(dirpath)
	Call AppCentre_BianLi(dirpath)

	Dim i,j,d,nd

	For Each i In AppCentre_LastModifiTime.Items

	  d=i
	  If nd="" Then nd=d
	  If DateDiff("s",nd,d)>0 Then nd=d

	Next

	AppCentre_GetLastModifiTime=Year(nd) &"-"&Month(nd)&"-"&Day(nd)
End Function

Dim AppCentre_LastModifiTime
Set AppCentre_LastModifiTime=CreateObject("Scripting.Dictionary")

Function AppCentre_BianLi(Path) '遍历递归搜索所有文件
	Dim Fso,ObjFolder,ObjFile 'Fso对象,子目录对象,文件对象
	Set Fso=Server.CreateObject("scripting.filesystemObject") '创建FSO读写对象
	For Each ObjFile in Fso.GetFolder(Path).Files '读取当前目录下的文件
		AppCentre_LastModifiTime.Add AppCentre_LastModifiTime.Count,ObjFile.DateLastModified
	Next
	For Each ObjFolder In Fso.GetFolder(Path).SubFolders '读取子目录
		AppCentre_BianLi(Path & "\" & ObjFolder.Name) '调用递归搜索子目录完整路径
	Next
End Function






'*********************************************************
Function InstallApp(FilePath)
	On Error Resume Next
	InstallApp=False
	Dim Install_Error
	Dim Install_Path
	Dim objXmlFile
	Dim objNodeList
	Dim objFSO
	Dim objStream
	Dim i,j

	Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")
	objXmlFile.async = False
	objXmlFile.ValidateOnParse=False
	objXmlFile.load(FilePath)
		
	If objXmlFile.readyState<>4 Then
		SetBlogHint_Custom "无法加载此文件！"
		Exit Function
	Else
		If objXmlFile.parseError.errorCode <> 0 Then
			SetBlogHint_Custom "该文件("&FilePath&")存在错误"
			Exit Function
		Else

			Dim Pack_ver,Pack_ID,Pack_Name
			Pack_Ver = objXmlFile.documentElement.SelectSingleNode("//app").getAttributeNode("version").value
			Pack_Type = objXmlFile.documentElement.selectSingleNode("//app").getAttributeNode("type").value
			Pack_For = objXmlFile.documentElement.selectSingleNode("//app").getAttributeNode("for").value
			app_adapted = objXmlFile.documentElement.selectSingleNode("//app").selectSingleNode("adapted").text

			If IsNumeric(app_adapted) Then
				If CLng(app_adapted)>CLng(BlogVersion) Then
					SetBlogHint_Custom "您的Z-Blog版本太低，无法安装该应用！"
					SetBlogHint_Custom "该应用需求Z-Blog版本：Z-Blog 2.0 Build " & app_adapted 
					SetBlogHint_Custom "您的Z-Blog版本：Z-Blog 2.0 Build " & BlogVersion
					Response.Redirect BlogHost & "zb_system/cmd.asp?act=PlugInMng"
					Exit Function
				End If
			Else
				SetBlogHint_Custom "该应用是为较低版本Z-Blog编写的应用，请仔细查看发布说明。"
			End If
			Pack_ID = objXmlFile.documentElement.selectSingleNode("id").text
			Pack_Name = objXmlFile.documentElement.selectSingleNode("name").text

			'If (CDbl(Pack_Ver) > CDbl(XML_Pack_Ver)) Then
			'	Response.Write "<p><font color=""red""> × ZPI 文件的 XML 版本为 "& Pack_Ver &", 而你的解包器版本为 "& XML_Pack_Ver &", 请升级您的 PluginSapper, 安装被中止.</font></p>"
			'	Exit Sub
			'ElseIf (LCase(Pack_Type) <> LCase(XML_Pack_Type)) Then
			'	Response.Write "<p><font color=""red""> × 不是 ZPI 文件, 而可能是 "& Pack_Type &", 安装被中止.</font></p>"
			'	Exit Sub
			'ElseIf (LCase(Pack_For) <> LCase(XML_Pack_Version)) Then
			'	Response.Write "<p><font color=""red""> × ZPI 文件版本不符合, 该版本可能是 "& Pack_For &", 安装被中止.</font></p>"
			'	Exit Sub
			'Else

			Install_Path=BlogPath & "zb_users\" & Pack_Type & "\"


			Set objNodeList = objXmlFile.documentElement.selectNodes("//folder/path")
			Set objFSO = CreateObject("Scripting.FileSystemObject")
				
				j=objNodeList.length-1
				For i=0 To j
					If objFSO.FolderExists(Install_Path & objNodeList(i).text)=False Then
						objFSO.CreateFolder(Install_Path & objNodeList(i).text)
					End If
				Next
			Set objFSO = Nothing
			Set objNodeList = Nothing
			Set objNodeList = objXmlFile.documentElement.selectNodes("//file/path")
			
				j=objNodeList.length-1
				For i=0 To j
					Set objStream = CreateObject("ADODB.Stream")
						With objStream
							.Type = 1
							.Open
							.Write objNodeList(i).nextSibling.nodeTypedvalue
							.SaveToFile Install_Path & objNodeList(i).text,2
							.Close
						End With
					Set objStream = Nothing
				Next
			Set objNodeList = Nothing

			'End If

			Call SetBlogHint_Custom("安装'<b>"& Pack_Name &" ("&Pack_ID&")</b>'成功!")
			InstallApp=True
		End If
	End If
		
	Set objXmlFile = Nothing


End Function
'*********************************************************




'*********************************************************
Function LoadPluginXmlInfo(id)

	On Error Resume Next

	app_path=BlogPath & "zb_users\plugin" & "\" & id & "\"

	Dim objXmlFile,strXmlFile
	Dim fso, s
	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FileExists(BlogPath & "zb_users/plugin" & "/" & id & "/" & "plugin.xml") Then

		strXmlFile =BlogPath & "zb_users/plugin" & "/" & id & "/" & "plugin.xml"

		Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		objXmlFile.load(strXmlFile)
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode <> 0 Then
			Else

				app_id=id
				app_name=objXmlFile.documentElement.selectSingleNode("name").text
				app_url=objXmlFile.documentElement.selectSingleNode("url").text

				app_adapted=objXmlFile.documentElement.selectSingleNode("adapted").text
				app_version=objXmlFile.documentElement.selectSingleNode("version").text
				app_pubdate=objXmlFile.documentElement.selectSingleNode("pubdate").text
				app_modified=objXmlFile.documentElement.selectSingleNode("modified").text

				app_note=objXmlFile.documentElement.selectSingleNode("note").text
				app_description=objXmlFile.documentElement.selectSingleNode("description").text

				app_author_name=objXmlFile.documentElement.selectSingleNode("author/name").text
				app_author_email=objXmlFile.documentElement.selectSingleNode("author/email").text
				app_author_url=objXmlFile.documentElement.selectSingleNode("author/url").text

				'app_source_name=objXmlFile.documentElement.selectSingleNode("source/name").text
				'app_source_email=objXmlFile.documentElement.selectSingleNode("source/email").text
				'app_source_url=objXmlFile.documentElement.selectSingleNode("source/url").text


				app_plugin_path=objXmlFile.documentElement.selectSingleNode("path").text
				app_plugin_include=objXmlFile.documentElement.selectSingleNode("include").text
				app_plugin_level=objXmlFile.documentElement.selectSingleNode("level").text

				app_price=objXmlFile.documentElement.selectSingleNode("price").text
				
				app_dependency=objXmlFile.documentElement.selectSingleNode("advanced/dependency").text
				app_rewritefunctions=objXmlFile.documentElement.selectSingleNode("advanced/rewritefunctions").text
				app_conflict=objXmlFile.documentElement.selectSingleNode("advanced/conflict").text

			End If
		End If
	End If

End Function
'*********************************************************





'*********************************************************
Function LoadThemeXmlInfo(id)

	On Error Resume Next

	Dim objXmlFile,strXmlFile
	Dim fso, s
	Set fso = CreateObject("Scripting.FileSystemObject")

	app_path=BlogPath & "zb_users\theme" & "\" & id & "\"

	If fso.FileExists(BlogPath & "zb_users/theme" & "/" & id & "/" & "theme.xml") Then

		strXmlFile =BlogPath & "zb_users/theme" & "/" & id & "/" & "theme.xml"

		Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		objXmlFile.load(strXmlFile)
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode <> 0 Then
			Else

				app_id=id
				app_name=objXmlFile.documentElement.selectSingleNode("name").text
				app_url=objXmlFile.documentElement.selectSingleNode("url").text

				app_adapted=objXmlFile.documentElement.selectSingleNode("adapted").text
				app_version=objXmlFile.documentElement.selectSingleNode("version").text
				app_pubdate=objXmlFile.documentElement.selectSingleNode("pubdate").text
				app_modified=objXmlFile.documentElement.selectSingleNode("modified").text

				app_note=objXmlFile.documentElement.selectSingleNode("note").text
				app_description=objXmlFile.documentElement.selectSingleNode("description").text

				app_author_name=objXmlFile.documentElement.selectSingleNode("author/name").text
				app_author_email=objXmlFile.documentElement.selectSingleNode("author/email").text
				app_author_url=objXmlFile.documentElement.selectSingleNode("author/url").text

				app_source_name=objXmlFile.documentElement.selectSingleNode("source/name").text
				app_source_email=objXmlFile.documentElement.selectSingleNode("source/email").text
				app_source_url=objXmlFile.documentElement.selectSingleNode("source/url").text

				app_plugin_path=objXmlFile.documentElement.selectSingleNode("plugin/path").text
				app_plugin_include=objXmlFile.documentElement.selectSingleNode("plugin/include").text
				app_plugin_level=objXmlFile.documentElement.selectSingleNode("plugin/level").text

				app_price=objXmlFile.documentElement.selectSingleNode("price").text

				app_dependency=objXmlFile.documentElement.selectSingleNode("advanced/dependency").text
				app_rewritefunctions=objXmlFile.documentElement.selectSingleNode("advanced/rewritefunctions").text
				app_conflict=objXmlFile.documentElement.selectSingleNode("advanced/conflict").text

				app_sidebars=objXmlFile.documentElement.selectSingleNode("sidebars").xml

				Dim objXmlitems,item,c
				Set objXmlitems=objXmlFile.documentElement.SelectNodes("functions/function")
				for each item in objXmlitems
					ExecuteGlobal "Dim app_function_" & item.getAttribute("filename")
					c=item.text
					Execute "app_function_" & item.getAttribute("filename") & "=c"
				Next
			End If
		End If
	End If

End Function
'*********************************************************





'*********************************************************
'创建一个空的XML文件，为写入文件作准备
Function CreatePluginXml(FilePath)
On Error Resume Next
	Dim Theme_Id
	Dim Theme_Name
	Dim Theme_Url
	Dim Theme_Note
	Dim Theme_Description
	Dim Theme_Pubdate
	Dim Theme_Source_Name
	Dim Theme_Source_Url
	Dim Theme_Source_Email
	Dim Theme_Author_Name
	Dim Theme_Author_Url
	Dim Theme_Author_Email
	Dim Theme_ScreenShot
	Dim Theme_Style_Name
	Dim Theme_Modified
	Dim Theme_Adapted
	Dim Theme_Version
	Dim Theme_Price
	Dim Theme_dependency
	Dim Theme_rewritefunctions
	Dim Theme_conflict

	Dim fso
	Dim strXmlFile
	Dim objXmlFile

	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(BlogPath & "zb_users/plugin" & "/" & ID & "/" & "plugin.xml") Then

		strXmlFile =BlogPath & "zb_users/plugin" & "/" & ID & "/" & "plugin.xml"

		Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		objXmlFile.load(strXmlFile)
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode <> 0 Then
			Else
				'objXmlFile.documentElement.selectSingleNode("adapted").text=BlogVersion
				'objXmlFile.Save BlogPath & "zb_users/plugin" & "/" & ID & "/" & "plugin.xml"

				Theme_Id=""
				Theme_Name=""
				Theme_Url=""
				Theme_Note=""
				Theme_Description=""
				Theme_Pubdate=""
				Theme_Source_Name=""
				Theme_Source_Url=""
				Theme_Author_Name=""
				Theme_Author_Url=""
				Theme_ScreenShot=""
				Theme_Style_Name=""
				Theme_dependency=""
				Theme_conflict=""
				Theme_rewritefunctions=""

				'Theme_Source_Name=objXmlFile.documentElement.selectSingleNode("source/name").text
				'Theme_Source_Url=objXmlFile.documentElement.selectSingleNode("source/url").text
				'Theme_Source_Email=objXmlFile.documentElement.selectSingleNode("source/email").text

				Theme_Author_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
				Theme_Author_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
				Theme_Author_Email=objXmlFile.documentElement.selectSingleNode("author/email").text


				If Theme_Author_Name="" Then
					Theme_Author_Name=Theme_Source_Name
					Theme_Author_Url=Theme_Source_Url
					Theme_Author_Email=Theme_Source_Email
				End If

				Theme_Id=ID
				'Theme_Id=objXmlFile.documentElement.selectSingleNode("id").text
				Theme_Name=objXmlFile.documentElement.selectSingleNode("name").text
				Theme_Url=objXmlFile.documentElement.selectSingleNode("url").text
				Theme_Note=objXmlFile.documentElement.selectSingleNode("note").text
				Theme_Pubdate=objXmlFile.documentElement.selectSingleNode("pubdate").text
				Theme_Modified=objXmlFile.documentElement.selectSingleNode("modified").text
				Theme_Description=objXmlFile.documentElement.selectSingleNode("description").text
				Theme_Version=objXmlFile.documentElement.selectSingleNode("version").text
				Theme_Adapted=objXmlFile.documentElement.selectSingleNode("adapted").text
				Theme_Price=objXmlFile.documentElement.selectSingleNode("price").text
				Theme_dependency=objXmlFile.documentElement.selectSingleNode("advanced/dependency").text
				Theme_rewritefunctions=objXmlFile.documentElement.selectSingleNode("advanced/rewritefunctions").text
				Theme_conflict=objXmlFile.documentElement.selectSingleNode("advanced/conflict").text
			End If
		End If
	End If

	Dim Plugin_ID,Plugin_Name,Plugin_URL,Plugin_PubDate,Plugin_Adapted,Plugin_Version,Plugin_Modified,Plugin_Note,Plugin_Description,Plugin_Price


	'程序开始执行时间
	Dim XmlDoc,Root,xRoot
	Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
		XmlDoc.async = False
		XmlDoc.ValidateOnParse=False
		Set Root = XmlDoc.createProcessingInstruction("xml","version='1.0' encoding='utf-8'")
		XmlDoc.appendChild(Root)
		Set xRoot = XmlDoc.appendChild(XmlDoc.CreateElement("app"))
			xRoot.setAttribute "version","2.0"
			xRoot.setAttribute "type","Plugin"
			xRoot.setAttribute "for","Z-Blog 2.0"
		Set xRoot = Nothing

		'写入文件信息
		Dim Author,AuthorName,AuthorURL,AuthorEmail
		Dim Advanced,Advanced_dependency,Advanced_rewritefunctions,Advanced_conflict


		Set Plugin_ID = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("id"))
			Plugin_ID.Text = Theme_Id
		Set Plugin_ID=Nothing

		Set Plugin_Name = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("name"))
			Plugin_Name.Text = Theme_Name
		Set Plugin_Name=Nothing

		Set Plugin_URL = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("url"))
			Plugin_URL.Text = Theme_Url
		Set Plugin_URL=Nothing

		Set Plugin_PubDate = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("pubdate"))
			Plugin_PubDate.Text = Theme_Pubdate
		Set Plugin_PubDate=Nothing

		Set Plugin_Modified = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("modified"))
			Plugin_Modified.Text = Theme_Modified
		Set Plugin_Modified=Nothing

		Set Plugin_Version = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("version"))
			Plugin_Version.Text = Theme_Version
		Set Plugin_Version=Nothing

		Set Plugin_Adapted = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("adapted"))
			Plugin_Adapted.Text = Theme_Adapted
		Set Plugin_Adapted=Nothing

		Set Plugin_Note = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("note"))
			Plugin_Note.Text = Theme_Note
		Set Plugin_Note=Nothing

		Set Plugin_Price = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("price"))
			Plugin_Price.Text = Theme_Price
		Set Plugin_Price=Nothing

		Dim CThemeDescription,XMLcdata
		Set Plugin_Description = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("description"))
			Set XMLcdata = XmlDoc.createNode("cdatasection", "","")
				XMLcdata.NodeValue = Theme_Description
			Set CThemeDescription = Plugin_Description.AppendChild(XMLcdata)
			Set CThemeDescription = Nothing
			Set Plugin_Description = Nothing
		Set Plugin_Description=Nothing


		Set Author = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("author"))

			Set AuthorName = Author.AppendChild(XmlDoc.CreateElement("name"))
				AuthorName.Text = Theme_Author_Name
			Set AuthorName=Nothing

			Set AuthorURL = Author.AppendChild(XmlDoc.CreateElement("url"))
				AuthorURL.Text = Theme_Author_Url
			Set AuthorURL=Nothing

			Set AuthorEmail = Author.AppendChild(XmlDoc.CreateElement("email"))
				AuthorEmail.Text = Theme_Author_Email
			Set AuthorEmail=Nothing

		Set Author=Nothing

		Set Advanced = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("advanced"))

			Set Advanced_dependency = Advanced.AppendChild(XmlDoc.CreateElement("name"))
				Advanced_dependency.Text = Theme_dependency
			Set Advanced_dependency=Nothing

			Set Advanced_rewritefunctions = Advanced.AppendChild(XmlDoc.CreateElement("url"))
				Advanced_rewritefunctions.Text = Theme_rewritefunctions
			Set Advanced_rewritefunctions=Nothing

			Set Advanced_conflict = Advanced.AppendChild(XmlDoc.CreateElement("email"))
				Advanced_conflict.Text = Theme_conflict
			Set Advanced_conflict=Nothing

		Set Advanced=Nothing

		XmlDoc.Save(FilePath)
		'Response.Write XmlDoc.Xml
		Set Root = Nothing
	Set XmlDoc = Nothing

End Function
'*********************************************************






'*********************************************************
'创建一个空的XML文件，为写入文件作准备
Function CreateThemeXml(FilePath)
On Error Resume Next

	Dim Theme_Id
	Dim Theme_Name
	Dim Theme_Url
	Dim Theme_Note
	Dim Theme_Description
	Dim Theme_Pubdate
	Dim Theme_Source_Name
	Dim Theme_Source_Url
	Dim Theme_Source_Email
	Dim Theme_Author_Name
	Dim Theme_Author_Url
	Dim Theme_Author_Email
	Dim Theme_ScreenShot
	Dim Theme_Style_Name
	Dim Theme_Modified
	Dim Theme_Adapted
	Dim Theme_Version
	Dim Theme_Price
	Dim Theme_dependency
	Dim Theme_rewritefunctions
	Dim Theme_conflict
	Dim fso
	Dim strXmlFile
	Dim objXmlFile

	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(BlogPath & "zb_users/theme" & "/" & ID & "/" & "theme.xml") Then

		strXmlFile =BlogPath & "zb_users/theme" & "/" & ID & "/" & "theme.xml"

		Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		objXmlFile.load(strXmlFile)
		If objXmlFile.readyState=4 Then
			If objXmlFile.parseError.errorCode <> 0 Then
			Else
				'objXmlFile.documentElement.selectSingleNode("adapted").text=BlogVersion
				'objXmlFile.Save BlogPath & "zb_users/theme" & "/" & ID & "/" & "theme.xml"
				
				Theme_Id=""
				Theme_Name=""
				Theme_Url=""
				Theme_Note=""
				Theme_Description=""
				Theme_Pubdate=""
				Theme_Source_Name=""
				Theme_Source_Url=""
				Theme_Author_Name=""
				Theme_Author_Url=""
				Theme_ScreenShot=""
				Theme_Style_Name=""
				Theme_dependency=""
				Theme_conflict=""
				Theme_rewritefunctions=""

				Theme_Source_Name=objXmlFile.documentElement.selectSingleNode("source/name").text
				Theme_Source_Url=objXmlFile.documentElement.selectSingleNode("source/url").text
				Theme_Source_Email=objXmlFile.documentElement.selectSingleNode("source/email").text


				Theme_Author_Name=objXmlFile.documentElement.selectSingleNode("author/name").text
				Theme_Author_Url=objXmlFile.documentElement.selectSingleNode("author/url").text
				Theme_Author_Email=objXmlFile.documentElement.selectSingleNode("author/email").text


				If Theme_Author_Name="" Then
					Theme_Author_Name=Theme_Source_Name
					Theme_Author_Url=Theme_Source_Url
					Theme_Author_Email=Theme_Source_Email
				End If

				Theme_Id=ID
				'Theme_Id=objXmlFile.documentElement.selectSingleNode("id").text
				Theme_Name=objXmlFile.documentElement.selectSingleNode("name").text
				Theme_Url=objXmlFile.documentElement.selectSingleNode("url").text
				Theme_Note=objXmlFile.documentElement.selectSingleNode("note").text
				Theme_Pubdate=objXmlFile.documentElement.selectSingleNode("pubdate").text
				Theme_Modified=objXmlFile.documentElement.selectSingleNode("modified").text
				Theme_Description=objXmlFile.documentElement.selectSingleNode("description").text
				Theme_Version=objXmlFile.documentElement.selectSingleNode("version").text
				Theme_Adapted=objXmlFile.documentElement.selectSingleNode("adapted").text
				Theme_Price=objXmlFile.documentElement.selectSingleNode("price").text
				Theme_dependency=objXmlFile.documentElement.selectSingleNode("advanced/dependency").text
				Theme_rewritefunctions=objXmlFile.documentElement.selectSingleNode("advanced/rewritefunctions").text
				Theme_conflict=objXmlFile.documentElement.selectSingleNode("advanced/conflict").text

			End If
		End If
	End If

	Dim Plugin_ID,Plugin_Name,Plugin_URL,Plugin_PubDate,Plugin_Adapted,Plugin_Version,Plugin_Modified,Plugin_Note,Plugin_Description,Plugin_Price

	'程序开始执行时间
	Dim XmlDoc,Root,xRoot
	Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
		XmlDoc.async = False
		XmlDoc.ValidateOnParse=False
		Set Root = XmlDoc.createProcessingInstruction("xml","version='1.0' encoding='utf-8'")
		XmlDoc.appendChild(Root)
		Set xRoot = XmlDoc.appendChild(XmlDoc.CreateElement("app"))
			xRoot.setAttribute "version","2.0"
			xRoot.setAttribute "type","Theme"
			xRoot.setAttribute "for","Z-Blog 2.0"
		Set xRoot = Nothing

		'写入文件信息
		Dim Author,AuthorName,AuthorURL,AuthorEmail

		Set Plugin_ID = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("id"))
			Plugin_ID.Text = Theme_Id
		Set Plugin_ID=Nothing

		Set Plugin_Name = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("name"))
			Plugin_Name.Text = Theme_Name
		Set Plugin_Name=Nothing

		Set Plugin_URL = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("url"))
			Plugin_URL.Text = Theme_Url
		Set Plugin_URL=Nothing

		Set Plugin_PubDate = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("pubdate"))
			Plugin_PubDate.Text = Theme_Pubdate
		Set Plugin_PubDate=Nothing

		Set Plugin_Modified = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("modified"))
			Plugin_Modified.Text = Theme_Modified
		Set Plugin_Modified=Nothing

		Set Plugin_Version = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("version"))
			Plugin_Version.Text = Theme_Version
		Set Plugin_Version=Nothing

		Set Plugin_Adapted = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("adapted"))
			Plugin_Adapted.Text = Theme_Adapted
		Set Plugin_Adapted=Nothing

		Set Plugin_Note = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("note"))
			Plugin_Note.Text = Theme_Note
		Set Plugin_Note=Nothing

		Set Plugin_Price = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("price"))
			Plugin_Price.Text = Theme_Price
		Set Plugin_Price=Nothing


		Dim CThemeDescription,XMLcdata
		Set Plugin_Description = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("description"))
			Set XMLcdata = XmlDoc.createNode("cdatasection", "","")
				XMLcdata.NodeValue = Theme_Description
			Set CThemeDescription = Plugin_Description.AppendChild(XMLcdata)
			Set CThemeDescription = Nothing
			Set Plugin_Description = Nothing
		Set Plugin_Description=Nothing


		Set Author = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("author"))

			Set AuthorName = Author.AppendChild(XmlDoc.CreateElement("name"))
				AuthorName.Text = Theme_Author_Name
			Set AuthorName=Nothing

			Set AuthorURL = Author.AppendChild(XmlDoc.CreateElement("url"))
				AuthorURL.Text = Theme_Author_Url
			Set AuthorURL=Nothing

			Set AuthorEmail = Author.AppendChild(XmlDoc.CreateElement("email"))
				AuthorEmail.Text = Theme_Author_Email
			Set AuthorEmail=Nothing

		Set Author=Nothing
		
		Set Advanced = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("advanced"))

			Set Advanced_dependency = Advanced.AppendChild(XmlDoc.CreateElement("name"))
				Advanced_dependency.Text = Theme_dependency
			Set Advanced_dependency=Nothing

			Set Advanced_rewritefunctions = Advanced.AppendChild(XmlDoc.CreateElement("url"))
				Advanced_rewritefunctions.Text = Theme_rewritefunctions
			Set Advanced_rewritefunctions=Nothing

			Set Advanced_conflict = Advanced.AppendChild(XmlDoc.CreateElement("email"))
				Advanced_conflict.Text = Theme_conflict
			Set Advanced_conflict=Nothing

		Set Advanced=Nothing


		XmlDoc.Save(FilePath)
		Set Root = Nothing
	Set XmlDoc = Nothing

End Function
'*********************************************************





'*********************************************************
'遍历目录内的所有文件以及文件夹
Function LoadAppFiles(DirPath,FilePath)
'On Error Resume Next

	Dim XmlDoc
	Dim fso            'fso对象
	Dim objFolder      '文件夹对象
	Dim objSubFolders  '子文件夹集合
	Dim objSubFolder   '子文件夹对象
	Dim objFiles       '文件集合
	Dim objFile        '文件对象
	Dim objStream
	Dim pathname,TextStream,pp,Xfolder,Xfpath,Xfile,Xpath,Xstream
	Dim PathNameStr

	Set fso=server.CreateObject("scripting.filesystemobject")
	Set objFolder=fso.GetFolder(DirPath)'创建文件夹对象
	

	Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
	XmlDoc.async = False
	XmlDoc.ValidateOnParse=False
	XmlDoc.load (FilePath)

	'写入每个文件夹路径
	Set Xfolder = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("folder"))
	Set Xfpath = Xfolder.AppendChild(XmlDoc.CreateElement("path"))
		Xfpath.text = replace(DirPath,ZipPathDir,Pack_PluginDir)

		Set objFiles=objFolder.Files
			for each objFile in objFiles
				If lcase(DirPath & objFile.name) <> lcase(Request.ServerVariables("PATH_TRANSLATED")) Then
					PathNameStr = DirPath & "" & objFile.name
					'================================================
					'写入文件的路径及文件内容
				   Set Xfile = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("file"))
				   Set Xpath = Xfile.AppendChild(XmlDoc.CreateElement("path"))
					   Xpath.text = replace(PathNameStr,ZipPathDir,Pack_PluginDir)
				   '创建文件流读入文件内容，并写入XML文件中
				   Set objStream = Server.CreateObject("ADODB.Stream")
				   objStream.Type = 1
				   objStream.Open()
				   objStream.LoadFromFile(PathNameStr)
				   objStream.position = 0
				   
				   Set Xstream = Xfile.AppendChild(XmlDoc.CreateElement("stream"))
					   Xstream.SetAttribute "xmlns:dt","urn:schemas-microsoft-com:datatypes"
					   '文件内容采用二制方式存放
					   Xstream.dataType = "bin.base64"
					   Xstream.nodeTypedValue = objStream.Read()
				   
				   Set objStream=Nothing
				   Set Xpath = Nothing
				   Set Xstream = Nothing
				   Set Xfile = Nothing
				  '================================================
				end if
			next
	XmlDoc.Save(FilePath)
	Set Xfpath = Nothing
	Set Xfolder = Nothing
	Set XmlDoc = Nothing
	
	'创建的子文件夹对象
	Set objSubFolders=objFolder.Subfolders
		'调用递归遍历子文件夹
		for each objSubFolder in objSubFolders
			pathname = DirPath & objSubFolder.name & "\"
			Call LoadAppFiles(pathname,FilePath)
		next
	Set objFolder=Nothing
	Set objSubFolders=Nothing
	Set fso=Nothing

	'Err.Clear

End Function
'*********************************************************




'*********************************************************
Function SavePluginXmlInfo(id)

	Dim objXMLdoc
	Set objXMLdoc =Server.CreateObject("Microsoft.XMLDOM")

	Dim objPI,objXMLrss

	'Set objPI = objXMLdoc.createProcessingInstruction("xml","version=""1.0"" encoding=""utf-8"" standalone=""yes""")
	'objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
	'Set objPI = Nothing


	Set objXMLrss = objXMLdoc.createElement("plugin")

	objXMLdoc.AppendChild(objXMLrss)

	objXMLrss.setAttribute "version","2.0"



app_id=id'Request.Form("app_id")
app_name=Request.Form("app_name")
app_url=Request.Form("app_url")
app_note=Request.Form("app_note")

app_author_name=Request.Form("app_author_name")
app_author_email=Request.Form("app_author_email")
app_author_url=Request.Form("app_author_url")

app_source_name=Request.Form("app_source_name")
app_source_email=Request.Form("app_source_email")
app_source_url=Request.Form("app_source_url")

app_plugin_path=Request.Form("app_plugin_path")
app_plugin_include=Request.Form("app_plugin_include")
app_plugin_level=Request.Form("app_plugin_level")

app_adapted=Request.Form("app_adapted")
app_version=Request.Form("app_version")
app_pubdate=Request.Form("app_pubdate")
app_modified=Request.Form("app_modified")
app_description=Request.Form("app_description")
app_price=Request.Form("app_price")


app_dependency=Request.Form("app_dependency")
app_rewritefunctions=Request.Form("app_rewritefunctions")
app_conflict=Request.Form("app_conflict")




	Dim objXMLitem,objXMLcdata

	Set objXMLitem = objXMLdoc.createElement("id")
	objXMLitem.text=app_id
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("name")
	objXMLitem.text=app_name
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("url")
	objXMLitem.text=app_url
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("note")
	objXMLitem.text=app_note
	objXMLrss.AppendChild(objXMLitem)


	Set objXMLitem = objXMLdoc.createElement("path")
	objXMLitem.text=app_plugin_path
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("include")
	objXMLitem.text=app_plugin_include
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("level")
	objXMLitem.text=app_plugin_level
	objXMLrss.AppendChild(objXMLitem)


	Dim objXMLauthor
	Set objXMLauthor = objXMLdoc.createElement("author")

	Set objXMLitem = objXMLdoc.createElement("name")
	objXMLitem.text=app_author_name
	objXMLauthor.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("email")
	objXMLitem.text=app_author_email
	objXMLauthor.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("url")
	objXMLitem.text=app_author_url
	objXMLauthor.AppendChild(objXMLitem)
	

	objXMLrss.AppendChild(objXMLauthor)



	Set objXMLitem = objXMLdoc.createElement("adapted")
	objXMLitem.text=app_adapted
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("version")
	objXMLitem.text=app_version
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("pubdate")
	objXMLitem.text=app_pubdate
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("modified")
	objXMLitem.text=app_modified
	objXMLrss.AppendChild(objXMLitem)

	If app_description="" Then
		Set objXMLitem = objXMLdoc.createElement("description")
		objXMLitem.text=app_description
		objXMLrss.AppendChild(objXMLitem)
	Else
		objXMLrss.AppendChild(objXMLdoc.createElement("description"))
		Set objXMLcdata = objXMLdoc.createNode("cdatasection", "","")
		objXMLcdata.NodeValue=app_description
		objXMLrss.selectSingleNode("description").AppendChild(objXMLcdata)
	End If


	Set objXMLitem = objXMLdoc.createElement("price")
	objXMLitem.text=app_price
	objXMLrss.AppendChild(objXMLitem)


	Set objXMLauthor = objXMLdoc.createElement("advanced")

	Set objXMLitem = objXMLdoc.createElement("dependency")
	objXMLitem.text=app_dependency
	objXMLauthor.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("rewritefunctions")
	objXMLitem.text=app_rewritefunctions
	objXMLauthor.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("conflict")
	objXMLitem.text=app_conflict
	objXMLauthor.AppendChild(objXMLitem)
	objXMLrss.AppendChild(objXMLauthor)
	

	Dim xml
	xml="<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & objXMLdoc.xml

	Call SaveToFile(BlogPath & "zb_users/plugin/"&id&"/plugin.xml",xml,"utf-8",False)


	Call SetBlogHint_Custom("保存插件'<b>"& app_name &" ("&app_id&")</b>'成功!")

End Function
'*********************************************************








'*********************************************************
Function SaveThemeXmlInfo(id)

	Dim objXMLdoc
	Set objXMLdoc =Server.CreateObject("Microsoft.XMLDOM")

	Dim objPI,objXMLrss

	'Set objPI = objXMLdoc.createProcessingInstruction("xml","version=""1.0"" encoding=""utf-8"" standalone=""yes""")
	'objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
	'Set objPI = Nothing


	Set objXMLrss = objXMLdoc.createElement("theme")

	objXMLdoc.AppendChild(objXMLrss)

	objXMLrss.setAttribute "version","2.0"



app_id=id'Request.Form("app_id")
app_name=Request.Form("app_name")
app_url=Request.Form("app_url")
app_note=Request.Form("app_note")

app_author_name=Request.Form("app_author_name")
app_author_email=Request.Form("app_author_email")
app_author_url=Request.Form("app_author_url")

app_source_name=Request.Form("app_source_name")
app_source_email=Request.Form("app_source_email")
app_source_url=Request.Form("app_source_url")

app_plugin_path=Request.Form("app_plugin_path")
app_plugin_include=Request.Form("app_plugin_include")
app_plugin_level=Request.Form("app_plugin_level")

app_adapted=Request.Form("app_adapted")
app_version=Request.Form("app_version")
app_pubdate=Request.Form("app_pubdate")
app_modified=Request.Form("app_modified")
app_description=Request.Form("app_description")
app_price=Request.Form("app_price")


app_dependency=Request.Form("app_dependency")
app_rewritefunctions=Request.Form("app_rewritefunctions")
app_conflict=Request.Form("app_conflict")

app_sidebars=Request.Form("app_sidebars")


	Dim objXMLitem,objXMLcdata

	Set objXMLitem = objXMLdoc.createElement("id")
	objXMLitem.text=app_id
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("name")
	objXMLitem.text=app_name
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("url")
	objXMLitem.text=app_url
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("note")
	objXMLitem.text=app_note
	objXMLrss.AppendChild(objXMLitem)


	Dim objXMLauthor
	Set objXMLauthor = objXMLdoc.createElement("author")

	Set objXMLitem = objXMLdoc.createElement("name")
	objXMLitem.text=app_author_name
	objXMLauthor.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("email")
	objXMLitem.text=app_author_email
	objXMLauthor.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("url")
	objXMLitem.text=app_author_url
	objXMLauthor.AppendChild(objXMLitem)

	objXMLrss.AppendChild(objXMLauthor)


	Dim objXMLsource
	Set objXMLsource = objXMLdoc.createElement("source")

	Set objXMLitem = objXMLdoc.createElement("name")
	objXMLitem.text=app_source_name
	objXMLsource.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("email")
	objXMLitem.text=app_source_email
	objXMLsource.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("url")
	objXMLitem.text=app_source_url
	objXMLsource.AppendChild(objXMLitem)

	objXMLrss.AppendChild(objXMLsource)


	Dim objXMLplugin
	Set objXMLplugin = objXMLdoc.createElement("plugin")

	Set objXMLitem = objXMLdoc.createElement("path")
	objXMLitem.text=app_plugin_path
	objXMLplugin.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("include")
	objXMLitem.text=app_plugin_include
	objXMLplugin.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("level")
	objXMLitem.text=app_plugin_level
	objXMLplugin.AppendChild(objXMLitem)

	If app_plugin_include<>"" Or app_plugin_path<>"" Then
		objXMLrss.AppendChild(objXMLplugin)
	End If




	Set objXMLitem = objXMLdoc.createElement("adapted")
	objXMLitem.text=app_adapted
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("version")
	objXMLitem.text=app_version
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("pubdate")
	objXMLitem.text=app_pubdate
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("modified")
	objXMLitem.text=app_modified
	objXMLrss.AppendChild(objXMLitem)

	If app_description="" Then
		Set objXMLitem = objXMLdoc.createElement("description")
		objXMLitem.text=app_description
		objXMLrss.AppendChild(objXMLitem)
	Else
		objXMLrss.AppendChild(objXMLdoc.createElement("description"))
		Set objXMLcdata = objXMLdoc.createNode("cdatasection", "","")
		objXMLcdata.NodeValue=app_description
		objXMLrss.selectSingleNode("description").AppendChild(objXMLcdata)
	End If


	Set objXMLauthor = objXMLdoc.createElement("advanced")

	Set objXMLitem = objXMLdoc.createElement("dependency")
	objXMLitem.text=app_dependency
	objXMLauthor.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("rewritefunctions")
	objXMLitem.text=app_rewritefunctions
	objXMLauthor.AppendChild(objXMLitem)

	Set objXMLitem = objXMLdoc.createElement("conflict")
	objXMLitem.text=app_conflict
	objXMLauthor.AppendChild(objXMLitem)
	objXMLrss.AppendChild(objXMLauthor)
	
	
	Set objXMLitem = objXMLdoc.createElement("price")
	objXMLitem.text=app_price
	objXMLrss.AppendChild(objXMLitem)

	Set objXMLauthor = objXMLdoc.createElement("sidebars")

If InStr(app_sidebars,"sidebar1")>0 Then
	Set objXMLitem = objXMLdoc.createElement("sidebar1")
	objXMLitem.text=ZC_SIDEBAR_ORDER
	objXMLauthor.AppendChild(objXMLitem)
End If

If InStr(app_sidebars,"sidebar2")>0 Then
	Set objXMLitem = objXMLdoc.createElement("sidebar2")
	objXMLitem.text=ZC_SIDEBAR_ORDER2
	objXMLauthor.AppendChild(objXMLitem)
End If

If InStr(app_sidebars,"sidebar3")>0 Then
	Set objXMLitem = objXMLdoc.createElement("sidebar3")
	objXMLitem.text=ZC_SIDEBAR_ORDER3
	objXMLauthor.AppendChild(objXMLitem)
End If

If InStr(app_sidebars,"sidebar4")>0 Then
	Set objXMLitem = objXMLdoc.createElement("sidebar4")
	objXMLitem.text=ZC_SIDEBAR_ORDER4
	objXMLauthor.AppendChild(objXMLitem)
End If

If InStr(app_sidebars,"sidebar5")>0 Then
	Set objXMLitem = objXMLdoc.createElement("sidebar5")
	objXMLitem.text=ZC_SIDEBAR_ORDER5
	objXMLauthor.AppendChild(objXMLitem)
End If

	objXMLrss.AppendChild(objXMLauthor)


	Set objXMLauthor = objXMLdoc.createElement("functions")

Call GetFunction()
Dim fun
For Each fun In Functions
	If IsObject(fun)=True Then
		If fun.Source="theme_"&app_id Then
	Set objXMLitem = objXMLdoc.createElement("function")
	'objXMLitem.text="<name>"&TransferHTML(fun.Name,"[html-format]")&"</name><filename>"&TransferHTML(fun.FileName,"[html-format]")&"</filename><content>"&TransferHTML(fun.Content,"[html-format]")&"</content><htmlid>"&TransferHTML(fun.HtmlID,"[html-format]")&"</htmlid><ftype>"&fun.FType&"</ftype><maxli>"&fun.MaxLi&"</maxli>"
objXMLitem.setAttribute "name",TransferHTML(fun.Name,"[html-format]")
objXMLitem.setAttribute "filename",TransferHTML(fun.FileName,"[html-format]")
objXMLitem.setAttribute "htmlid",TransferHTML(fun.HtmlID,"[html-format]")
objXMLitem.setAttribute "ftype",TransferHTML(fun.FType,"[html-format]")
objXMLitem.setAttribute "maxli",TransferHTML(fun.MaxLi,"[html-format]")
objXMLitem.text=Replace(Request.Form("app_function_"&fun.filename),vbCrlf,"")'fun.Content

	objXMLauthor.AppendChild(objXMLitem)
		End If
	End If
Next


	objXMLrss.AppendChild(objXMLauthor)

	Dim xml
	xml="<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>" & objXMLdoc.xml

	Call SaveToFile(BlogPath & "zb_users/theme/"&id&"/theme.xml",xml,"utf-8",False)

	Call SetBlogHint_Custom("保存主题'<b>"& app_name &" ("&app_id&")</b>'成功!")

End Function
'*********************************************************



'*********************************************************
Function CreateNewPlugin(id)

	Dim strInclude,strMain,strFunction
	strFunction="<"&"%" & vbCrlf & _
				"'****************************************" & vbCrlf & _
				"' __name__ 子菜单"& vbCrlf & _
				"'****************************************"& vbCrlf & _
				"Function __name___SubMenu(id)" & vbCrlf & _
				"	Dim aryName,aryPath,aryFloat,aryInNewWindow,i" & vbCrlf & _
				"	aryName=Array(""首页"")" & vbCrlf & _
				"	aryPath=Array(""__path__"")" & vbCrlf & _
				"	aryFloat=Array(""m-left"")" & vbCrlf & _
				"	aryInNewWindow=Array(False)" & vbCrlf & _
				"	For i=0 To Ubound(aryName)" & vbCrlf & _
				"		__name___SubMenu=__name___SubMenu & MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id,"" m-now"",""""),aryInNewWindow(i))"& vbCrlf & _
				"	Next" & vbCrlf & _
				"End Function" & vbCrlf & _
				"%"&">"
	strInclude="<!-- #includ"&"e file=""function.asp"" -->" & vbCrlf & _
				"<"&"%" & vbCrlf & _
				"'注册插件" & vbCrlf & _
				"Call RegisterPlugin(""__name__"",""ActivePlugin___name__"")" & vbCrlf & _
				"'挂口部分" & vbCrlf & _
				"Function ActivePlugin___name__()" & vbCrlf & _
				"" & vbCrlf & _
				"	'插件最主要在这里挂接口。" & vbCrlf & _
				"	'Z-Blog可挂的接口有三类：Action、Filter、Response" & vbCrlf & _
				"	'建议参考Z-Wiki进行开发" & vbCrlf & _
				"	" & vbCrlf & _
				"End Function" & vbCrlf & vbCrlf & vbCrlf & _
				"Function InstallPlugin___name__()" & vbCrlf & _
				"" & vbCrlf & _
				"	'用户激活插件之后的操作" & vbCrlf & _
				"	" & vbCrlf & _
				"End Function" & vbCrlf & vbCrlf & vbCrlf & _
				"Function UnInstallPlugin___name__()" & vbCrlf & _
				"" & vbCrlf & _
				"	'用户停用插件之后的操作" & vbCrlf & _
				"	" & vbCrlf & _
				"End Function" & vbCrlf & _
				"%"&">"
	strMain="<"&"%@ LANGUAGE=""VBSCRIPT"" CODEPAGE=""65001""%"&">" & vbCrlf & _
			"<"&"% Option Explicit %"&">" & vbCrlf & _
			"<"&"% 'On Error Resume Next %"&">" & vbCrlf & _
			"<"&"% Response.Charset=""UTF-8"" %"&">" & vbCrlf & _
			"<!-- #inclu"&"de file=""..\..\c_option.asp"" -->" & vbCrlf & _
			"<!-- #inclu"&"de file=""..\..\..\zb_system\function\c_function.asp"" -->" & vbCrlf & _
			"<!-- #inclu"&"de file=""..\..\..\zb_system\function\c_system_lib.asp"" -->" & vbCrlf & _
			"<!-- #inclu"&"de file=""..\..\..\zb_system\function\c_system_base.asp"" -->" & vbCrlf & _
			"<!-- #inclu"&"de file=""..\..\..\zb_system\function\c_system_event.asp"" -->" & vbCrlf & _
			"<!-- #inclu"&"de file=""..\..\..\zb_system\function\c_system_manage.asp"" -->" & vbCrlf & _
			"<!-- #inclu"&"de file=""..\..\..\zb_system\function\c_system_plugin.asp"" -->" & vbCrlf & _
			"<!-- #inclu"&"de file=""..\p_config.asp"" -->" & vbCrlf & _
			"<"&"%" & vbCrlf & _
			"Call System_Initialize()" & vbCrlf & _
			"'检查非法链接" & vbCrlf & _
			"Call CheckReference("""")" & vbCrlf & _
			"'检查权限" & vbCrlf & _
			"If BlogUser.Level>__level__ Then Call ShowError(6)" & vbCrlf & _
			"If CheckPluginState(""__name__"")=False Then Call ShowError(48)" & vbCrlf & _
			"BlogTitle=""__title__""" & vbCrlf & _
			"%"&">" & vbCrlf & _
			"<!--#inclu"&"de file=""..\..\..\zb_system\admin\admin_header.asp""-->" & vbCrlf & _
			"<!--#inclu"&"de file=""..\..\..\zb_system\admin\admin_top.asp""-->" & vbCrlf & _
			"        <div id=""divMain"">" & vbCrlf & _
			"          <div id=""ShowBlogHint"">" & vbCrlf & _
			"            <"&"%Call GetBlogHint()%"&">" & vbCrlf & _
			"          </div>" & vbCrlf & _
			"          <div class=""divHeader""><"&"%=BlogTitle%"&"></div>" & vbCrlf & _
			"          <div class=""SubMenu""><"&"%=__name___SubMenu(0)%"&"></div>" & vbCrlf & _
			"          <div id=""divMain2""> " & vbCrlf & _
			"            <script type=""text/javascript"">ActiveTopMenu(""aPlugInMng"");</script> " & vbCrlf & _
			"            在这里写入后台管理页面代码" & vbCrlf & _
			"          </div>" & vbCrlf & _
			"        </div>" & vbCrlf & _
			"        <!--#inclu"&"de file=""..\..\..\zb_system\admin\admin_footer.asp""-->" & vbCrlf & _
			"" & vbCrlf & _
			"<"&"%Call System_Terminate()%"&">" & vbCrlf
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
		

	If objFSO.FolderExists(BlogPath & "zb_users\plugin\"&id)=False Then
		Call objFSO.CreateFolder(BlogPath & "zb_users\plugin\"&id)
		If Not IsEmpty(Request.Form("app_plugin_include")) Then
			strInclude=Replace(strInclude,"__name__",id)
			Call SaveToFile(BlogPath & "zb_users\plugin\"&id&"\"&Request.Form("app_plugin_include"),strInclude,"utf-8",False)
		End If
		If Not IsEmpty(Request.Form("app_plugin_path")) Then
			strMain=Replace(strMain,"__name__",id)
			strMain=Replace(strMain,"__level__",Request.Form("app_plugin_level"))
			strMain=Replace(strMain,"__title__",Request.Form("app_name"))
			strFunction=Replace(strFunction,"__path__",Request.Form("app_plugin_path"))
			strFunction=Replace(strFunction,"__name__",id)
			Call SaveToFile(BlogPath & "zb_users\plugin\"&id&"\"&Request.Form("app_plugin_path"),strMain,"utf-8",False)
			Call SaveToFile(BlogPath & "zb_users\plugin\"&id&"\function.asp",strFunction,"utf-8",False)
		End If
	Else
		Call SetBlogHint_Custom("已存在有相同ID的插件!!!")
		Response.Redirect Request.ServerVariables("HTTP_REFERER")
	End If

End Function
'*********************************************************



'*********************************************************
Function CreateNewTheme(id)

	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
		

	If objFSO.FolderExists(BlogPath & "zb_users\theme\"&id)=False Then
		Call objFSO.CreateFolder(BlogPath & "zb_users\theme\"&id)
	Else
		Call SetBlogHint_Custom("已存在有相同ID的主题!!!")
		Response.Redirect Request.ServerVariables("HTTP_REFERER")
	End If


End Function
'*********************************************************







'*********************************************************
Function ListTheme(s)
Response.Write "App应用中心筹备中,敬请期待!"
Exit Function
	Dim i,j
	Dim objXmlFile,objNodeList
	Set objXmlFile=ReturnXML(s)
	If objXmlFile.readyState=4 Then
		If objXmlFile.parseError.errorCode <> 0 Then
		Else
			Set objNodeList = objXmlFile.documentElement.selectNodes("//data/app")
				
			j=objNodeList.length-1
			For i=0 To j
				Response.Write "<div class='app-theme'>"
				Response.Write "<p class='app-theme-id'>" & objNodeList(i).SelectSingleNode("id").text & "</p>"
				Response.Write "<p class='app-theme-name'><a href='?act=detail&id=" & objNodeList(i).SelectSingleNode("id").text & "'>" & objNodeList(i).SelectSingleNode("name").text & "</a></p>"
				Response.Write "<p class='app-theme-image'><img width='320' height='240' src='images/theme.png' alt='' title='' /></p>"
				Response.Write "</div>"
			Next

		End If
	End If


End Function
'*********************************************************




'*********************************************************
Function ListPlugin(s)
Response.Write "App应用中心筹备中,敬请期待!"
Exit Function
	Dim i,j
	Dim objXmlFile,objNodeList
	Set objXmlFile=ReturnXML(s)
	If objXmlFile.readyState=4 Then
		If objXmlFile.parseError.errorCode <> 0 Then
		Else
			Set objNodeList = objXmlFile.documentElement.selectNodes("//data/app")
			j=objNodeList.length-1
			For i=0 To j
				Response.Write "<div class='app-plugin'>"
				Response.Write "<p class='app-plugin-id'>" & objNodeList(i).SelectSingleNode("id").text & "</p>"
				Response.Write "<p class='app-plugin-name'><a href='?act=detail&id=" & objNodeList(i).SelectSingleNode("id").text & "'>" & objNodeList(i).SelectSingleNode("name").text & "</a></p>"
				Response.Write "<p class='app-plugin-image'><img width='128' height='128' src='images/plugin.png' alt='' title='' /></p>"
				Response.Write "</div>"
			Next

		End If
	End If

End Function
'*********************************************************
'*********************************************************

Function DetailPlugin(s)

	Dim i,j
	Dim objXmlFile,objNodeList
	Set objXmlFile=ReturnXML(s)
	If objXmlFile.readyState=4 Then
		If objXmlFile.parseError.errorCode <> 0 Then
		Else
			Set objNodeList = objXmlFile.documentElement.selectNodes("//data/app")
			j=objNodeList.length-1
			For i=0 To j
				Response.Write "<div class='app-plugin'>"
				Response.Write "<p class='app-plugin-id'>" & objNodeList(i).SelectSingleNode("id").text & "</p>"
				Response.Write "<p class='app-plugin-name'><a href='?act=detail&id=" & objNodeList(i).SelectSingleNode("id").text & "'>" & objNodeList(i).SelectSingleNode("name").text & "</a></p>"
				Response.Write "<p class='app-plugin-image'><img width='128' height='128' src='images/plugin.png' alt='' title='' /></p>"
				Response.Write "<p class='app-plugin-lastupdate'>最后更新："&objNodeList(i).SelectSingleNode("lastupdate").text&"</p>"
				Response.Write "<p class='app-plugin-version'>版本号："&objNodeList(i).SelectSingleNode("version").text&"</p>"
				Response.Write "<p class='app-plugin-zbver'>支持Z-Blog版本："&objNodeList(i).SelectSingleNode("zbver").text&"</p>"
				Response.Write "<p class='app-plugin-size'>大小："&objNodeList(i).SelectSingleNode("size").text&"</p>"
				Response.Write "<p class='app-plugin-tag'>标签："&TagToUrl(objNodeList(i).SelectSingleNode("tag").text)&"</p>"
				Response.Write "<p class='app-plugin-pay'>金额：￥"&objNodeList(i).SelectSingleNode("pay").text&"</p>"

				Response.Write "<p class='app-plugin-author'>作者：<a href='"&objXmlFile.documentElement.selectNodes("//data/app/author")(0).SelectSingleNode("url").text&"' target='_blank'>"&objXmlFile.documentElement.selectNodes("//data/app/author")(0).SelectSingleNode("name").text&"</a></p>"

				Response.Write "<p class='app-plugin-detail'><span class='app-plugin-down'><a href='?act=download&url="&Server.URLEncode(objNodeList(i).SelectSingleNode("downurl").text)&"'>下载</a></span><span class='app-plugin-down'><a href='"&objNodeList(i).SelectSingleNode("posturl").text&"' target='_blank'>查看</a></span></p>"
				Response.Write "</div>"
			Next

		End If
	End If

End Function
'*********************************************************
'*********************************************************

Function DetailTheme(s)

	Dim i,j
	Dim objXmlFile,objNodeList
	Set objXmlFile=ReturnXML(s)
	If objXmlFile.readyState=4 Then
		If objXmlFile.parseError.errorCode <> 0 Then
		Else
			Set objNodeList = objXmlFile.documentElement.selectNodes("//data/app")
			j=objNodeList.length-1
			For i=0 To j
				Response.Write "<div class='app-theme'>"
				Response.Write "<p class='app-theme-id'>" & objNodeList(i).SelectSingleNode("id").text & "</p>"
				Response.Write "<p class='app-theme-name'><a href='?act=detail&id=" & objNodeList(i).SelectSingleNode("id").text & "'>" & objNodeList(i).SelectSingleNode("name").text & "</a></p>"
				Response.Write "<p class='app-theme-image'><img width='128' height='128' src='images/theme.png' alt='' title='' /></p>"
				Response.Write "<p class='app-theme-lastupdate'>最后更新："&objNodeList(i).SelectSingleNode("lastupdate").text&"</p>"
				Response.Write "<p class='app-theme-version'>版本号："&objNodeList(i).SelectSingleNode("version").text&"</p>"
				Response.Write "<p class='app-theme-zbver'>支持Z-Blog版本："&objNodeList(i).SelectSingleNode("zbver").text&"</p>"
				Response.Write "<p class='app-theme-size'>大小："&objNodeList(i).SelectSingleNode("size").text&"</p>"
				Response.Write "<p class='app-theme-tag'>标签："&TagToUrl(objNodeList(i).SelectSingleNode("tag").text)&"</p>"
				Response.Write "<p class='app-theme-pay'>金额：￥"&objNodeList(i).SelectSingleNode("pay").text&"</p>"

				Response.Write "<p class='app-theme-author'>作者：<a href='"&objXmlFile.documentElement.selectNodes("//data/app/author")(0).SelectSingleNode("url").text&"' target='_blank'>"&objXmlFile.documentElement.selectNodes("//data/app/author")(0).SelectSingleNode("name").text&"</a></p>"

				Response.Write "<p class='app-theme-detail'><span class='app-theme-down'><a href='?act=download&url="&Server.URLEncode(objNodeList(i).SelectSingleNode("downurl").text)&"'>下载</a></span><span class='app-theme-down'><a href='"&objNodeList(i).SelectSingleNode("posturl").text&"' target='_blank'>查看</a></span></p>"
				Response.Write "</div>"
			Next

		End If
	End If

End Function
'*********************************************************


'*********************************************************
Function ReturnXML(s)
	Dim objXmlFile
	Set objXmlFile=Server.CreateObject("Microsoft.XMLDOM")
	objXmlFile.async = False
	objXmlFile.ValidateOnParse=False
	objXmlFile.loadXML(s)
	Set ReturnXML=objXmlFile
End Function
'*********************************************************

'*********************************************************
Function TagToUrl(s)
	If Instr(s,",")<=0 Then TagToUrl=s:Exit Function
	Dim arySpt,strTmp
	arySpt=Split(s,",")
	Dim i
	For i=0 To Ubound(arySpt)
		strTmp=strTmp & "<a href='?act=tag&tag="&arySpt(i)&"'>"&arySpt(i)&"</a>" 
	Next
	TagToUrl=strTmp
End Function


'*********************************************************
Function CRC32(Path)

	On Error Resume Next
	Dim objAdo
	If IsObject(Path) Then	
		Set objAdo=Path
	Else
		Set objAdo = CreateObject("adodb.stream")
		objAdo.Open
		objAdo.Type = 1
		objAdo.LoadFromFile Path
    End If   
    Dim crc32Result
    Dim i
    Dim j
    Dim dwCrc
    Dim iLookup
    Dim Lb
    Dim Ub
            
     '常数
    Const Num0 = &H0
    Const Num1 = &H1
    Const Num2 = &H2
    Const Num8 = &H8
    Const Num255 = &HFF
    Const Num256 = &H100
    Const Num16777215 = &HFFFFFF
    Const dwPolynomial = &HEDB88320
    Const Num2147483647 = &H7FFFFFFF
    Const NumNegative1 = &HFFFFFFFF
    Const NumNegative2 = &HFFFFFFFE
    Const NumNegative256 = &HFFFFFF00
            
    'CRC32表
    Dim crc32Table(&HFF)
            
    '初始化CRC32表
    For i = Num0 To Num255
        dwCrc = i
        For j = Num8 To Num1 Step NumNegative1
            If (dwCrc And Num1) Then
                dwCrc = ((dwCrc And NumNegative2) \ Num2) And Num2147483647
                dwCrc = dwCrc Xor dwPolynomial
            Else
                dwCrc = ((dwCrc And NumNegative2) \ Num2) And Num2147483647
            End If
        Next
        crc32Table(i) = dwCrc
    Next
    crc32Result = NumNegative1
    '计算CRC32码
    For i = 0 To objAdo.Size-1
        objAdo.Position=i
        iLookup = (crc32Result And Num255) Xor AscB(objAdo.Read(1))
        crc32Result = ((crc32Result And NumNegative256) \ Num256) And Num16777215
        crc32Result = crc32Result Xor crc32Table(iLookup)
    Next
    objAdo.Position=0
    CRC32 = Hex(Not (crc32Result))
End Function
'*********************************************************


%>