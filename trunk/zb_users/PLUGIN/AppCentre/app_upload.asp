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

Dim ZipPathFile
ZipPathFile=BlogPath & "zb_users\cache\temp.zba"


Dim objUpLoadFile
Set objUpLoadFile=New TUpLoadFile


objUpLoadFile.AutoName=False
objUpLoadFile.IsManual=True
objUpLoadFile.FileSize=0
objUpLoadFile.FileName="temp.zba"
objUpLoadFile.FullPath=ZipPathFile

If objUpLoadFile.UpLoad_Form() Then
	If objUpLoadFile.SaveFile() Then
		Call Install_Plugin(ZipPathFile)
	End If
End If






Response.Redirect Request.ServerVariables("HTTP_REFERER")



Sub Install_Plugin(FilePath)
'On Error Resume Next

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
	Else
		If objXmlFile.parseError.errorCode <> 0 Then
		Else

			Dim Pack_ver,Pack_Type,Pack_For,Pack_ID,Pack_Name
			Pack_Ver = objXmlFile.documentElement.SelectSingleNode("//app").getAttributeNode("version").value
			Pack_Type = objXmlFile.documentElement.selectSingleNode("//app").getAttributeNode("type").value
			Pack_For = objXmlFile.documentElement.selectSingleNode("//app").getAttributeNode("for").value
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

		End If
	End If
		
	Set objXmlFile = Nothing


End Sub
'*********************************************************

%>