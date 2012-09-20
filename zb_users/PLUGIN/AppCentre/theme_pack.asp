<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("AppCentre")=False Then Call ShowError(48)



Dim ZipPathDir,ZipPathFile,Pack_PluginDir,ID

ID=Request.QueryString("id")

ZipPathDir = BlogPath & "zb_users\theme\" & ID & "\"
ZipPathFile = BlogPath & "zb_users\cache\" & MD5(ZC_BLOG_CLSID & ID) & ".zba"
Pack_PluginDir = ID & "\"

Call CreateXml(ZipPathFile)




Response.AddHeader   "Content-Disposition","attachment;filename="& ID &".zba"
Response.ContentType = "application/octet-stream"

Server.Transfer "../../cache/"& MD5(ZC_BLOG_CLSID & ID) &".zba"






'遍历目录内的所有文件以及文件夹
Sub LoadData(DirPath)
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
	
	'Response.Write("<p>打包目录: "& Replace(DirPath,BlogPath,"") &"...</p>")
	'Response.flush
	
	Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
	XmlDoc.async = False
	XmlDoc.ValidateOnParse=False
	XmlDoc.load (ZipPathFile)

	'写入每个文件夹路径
	Set Xfolder = XmlDoc.SelectSingleNode("//app").AppendChild(XmlDoc.CreateElement("folder"))
	Set Xfpath = Xfolder.AppendChild(XmlDoc.CreateElement("path"))
		Xfpath.text = replace(DirPath,ZipPathDir,Pack_PluginDir)

	'Response.Write "<blockquote><font color=""Teal"">"
		Set objFiles=objFolder.Files
			for each objFile in objFiles
				If lcase(DirPath & objFile.name) <> lcase(Request.ServerVariables("PATH_TRANSLATED")) Then
					PathNameStr = DirPath & "" & objFile.name
					'Response.Write Replace(PathNameStr,BlogPath,"") & "<br />"
					'Response.flush
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
	'Response.Write "</font></blockquote>"
	XmlDoc.Save(ZipPathFile)
	Set Xfpath = Nothing
	Set Xfolder = Nothing
	Set XmlDoc = Nothing
	
	'创建的子文件夹对象
	Set objSubFolders=objFolder.Subfolders
		'调用递归遍历子文件夹
		for each objSubFolder in objSubFolders
			pathname = DirPath & objSubFolder.name & "\"
			LoadData(pathname)
		next
	Set objFolder=Nothing
	Set objSubFolders=Nothing
	Set fso=Nothing

	If Err.Number<>0 Then Pack_Error=Pack_Error+1
	Err.Clear

End Sub






'创建一个空的XML文件，为写入文件作准备
Sub CreateXml(FilePath)
'On Error Resume Next

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
				Theme_ScreenShot="../../zb_users/theme" &"/" & Theme_Id & "/" & "screenshot.png"
			End If
		End If
	End If

	Dim Plugin_ID,Plugin_Name,Plugin_URL,Plugin_PubDate,Plugin_Adapted,Plugin_Version,Plugin_Modified,Plugin_Note,Plugin_Description

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

		XmlDoc.Save(FilePath)
		Set Root = Nothing
	Set XmlDoc = Nothing

	LoadData(ZipPathDir)

End Sub


%>