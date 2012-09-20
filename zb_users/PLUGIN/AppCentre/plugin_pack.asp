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
<!-- #include file="function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("AppCentre")=False Then Call ShowError(48)



Dim ZipPathDir,ZipPathFile,Pack_PluginDir

ZipPathDir = BlogPath & "zb_users\plugin\" & Request.QueryString("id") & "\"
ZipPathFile = BlogPath & "zb_users\cache\" & Request.QueryString("id") & ".zba"
Pack_PluginDir = Request.QueryString("id") & "\"

Call CreateXml(ZipPathFile)


Response.AddHeader   "Content-Disposition","attachment;filename="& Request.QueryString("id") &".zba"
Response.ContentType = "application/octet-stream"

Server.Transfer "../../cache/"& Request.QueryString("id") &".zpi"


'遍历目录内的所有文件以及文件夹
Sub LoadData(DirPath)
On Error Resume Next

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
	Set Xfolder = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("folder"))
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
				   Set Xfile = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("file"))
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
On Error Resume Next

	'程序开始执行时间
	Dim XmlDoc,Root,xRoot
	Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
		XmlDoc.async = False
		XmlDoc.ValidateOnParse=False
		Set Root = XmlDoc.createProcessingInstruction("xml","version='1.0' encoding='utf-8'")
		XmlDoc.appendChild(Root)
		Set xRoot = XmlDoc.appendChild(XmlDoc.CreateElement("root"))
			xRoot.setAttribute "version","1.0"
			xRoot.setAttribute "type","Plugin"
			xRoot.setAttribute "for","Z-Blog_2_0"
		Set xRoot = Nothing

		'写入文件信息
		Dim Author,AuthorName,AuthorURL,AuthorEmail

		Set Plugin_ID = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("id"))
			Plugin_ID.Text = Request.Form("PluginID")
		Set Plugin_ID=Nothing

		Set Plugin_Name = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("name"))
			Plugin_Name.Text = Request.Form("PluginName")
		Set Plugin_Name=Nothing

		Set Plugin_URL = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("url"))
			Plugin_URL.Text = Request.Form("PluginURL")
		Set Plugin_URL=Nothing

		Set Plugin_PubDate = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("pubdate"))
			Plugin_PubDate.Text = Request.Form("PluginPubDate")
		Set Plugin_PubDate=Nothing

		Set Plugin_Adapted = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("adapted"))
			Plugin_Adapted.Text = Request.Form("PluginAdapted")
		Set Plugin_Adapted=Nothing

		Set Plugin_Version = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("version"))
			Plugin_Version.Text = Request.Form("PluginVersion")
		Set Plugin_Version=Nothing

		Set Plugin_Modified = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("modified"))
			Plugin_Modified.Text = Request.Form("PluginModified")
		Set Plugin_Modified=Nothing

		Set Plugin_Note = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("note"))
			Plugin_Note.Text = Replace(Replace(Request.Form("PluginNote"),vbCr,""),vbLf,"")
		Set Plugin_Note=Nothing

		Set Author = XmlDoc.SelectSingleNode("//root").AppendChild(XmlDoc.CreateElement("author"))

			Set AuthorName = Author.AppendChild(XmlDoc.CreateElement("name"))
				AuthorName.Text = Request.Form("AuthorName")
			Set AuthorName=Nothing

			Set AuthorURL = Author.AppendChild(XmlDoc.CreateElement("url"))
				AuthorURL.Text = Request.Form("AuthorURL")
			Set AuthorURL=Nothing

			Set AuthorEmail = Author.AppendChild(XmlDoc.CreateElement("email"))
				AuthorEmail.Text = Request.Form("AuthorEmail")
			Set AuthorEmail=Nothing

		Set Author=Nothing

		XmlDoc.Save(FilePath)
		'Response.Write XmlDoc.Xml
		Set Root = Nothing
	Set XmlDoc = Nothing

	LoadData(ZipPathDir)

End Sub





'Response.Write 123



%>