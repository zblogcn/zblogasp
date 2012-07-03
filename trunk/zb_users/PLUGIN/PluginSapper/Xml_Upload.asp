<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Pre Terminator 及以上版本, 其它版本的Z-blog未知
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    插件管理插件
'// 最后修改：   2008-6-28
'// 最后版本:    1.2
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Server.ScriptTimeout=99999999 %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="c_sapper.asp" -->
<!-- #include file="../p_config.asp" -->
<%

Call System_Initialize()

'检查非法链接
Call CheckReference("")

'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("PluginSapper")=False Then Call ShowError(48)

BlogTitle = "从本地导入插件"

%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
	<meta name="robots" content="noindex,nofollow"/>
	<link rel="stylesheet" rev="stylesheet" href="../../../ZB_SYSTEM/CSS/admin.css" type="text/css" media="screen" />
	<link rel="stylesheet" rev="stylesheet" href="images/style.css" type="text/css" media="screen" />
	<title><%=BlogTitle%></title>
</head>
<body>

<div id="divMain">
	<div class="Header">Plugin Sapper - 从本地上传 ZPI 文件并导入插件到 Blog. <a href="help.asp#importzpi"  title="关于导入插件">[页面帮助]</a></div>
	<%Call SapperMenu("3")%>
<div id="divMain2">
<%

'*********************************************************
' 目的：    定义TUpLoadFile类
' 输入：    无
' 返回：    无
'*********************************************************
Class TUpLoadFile2

	Public ID
	Public AuthorID

	Public FileSize
	Public FileName
	Public PostTime
	Public Stream

	Private FUploadType
	Public Property Let UploadType(strUploadType)
		If (strUploadType="Stream") Then
			FUploadType=strUploadType
		Else
			FUploadType="Form"
		End If
	End Property
	Public Property Get UploadType
		If IsEmpty(FUploadType)=True Then
			UploadType="Form"
		Else
			UploadType = FUploadType
		End If
	End Property

	Function UpLoad_Form()

		Dim i,j
		Dim x,y,z
		Dim intFormSize
		Dim binFormData
		Dim strFileName

		intFormSize = Request.TotalBytes
		binFormData = Request.BinaryRead(intFormSize)

		If Instr(CStr(Request.ServerVariables("HTTP_USER_AGENT")),"Opera")>0 Then
			i=0
			i=InstrB(binFormData,ChrB(13)&ChrB(10)&ChrB(13)&ChrB(10))
			If i>0 Then i=i+3
			j=InstrB(binFormData,ChrB(13)&ChrB(10)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45))
		ElseIf  Instr(CStr(Request.ServerVariables("HTTP_USER_AGENT")),"AppleWebKit")>0 Then
			i=0
			i=InstrB(binFormData,ChrB(13)&ChrB(10)&ChrB(13)&ChrB(10))
			If i>0 Then i=i+3
			j=InstrB(binFormData,ChrB(13)&ChrB(10)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45))
		Else
			i=InstrB(binFormData,ChrB(13)&ChrB(10)&ChrB(13)&ChrB(10))
			i=i+3
			j=InStrB(binFormData,ChrB(13)&ChrB(10)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45)&ChrB(45))
		End If 


		If Len(Request.QueryString("filename"))>0 Then
			strFileName=Request.QueryString("filename")
		Else
			x=InstrB(binFormData,ChrB(&H66)&ChrB(&H69)&ChrB(&H6C)&ChrB(&H65)&ChrB(&H6E)&ChrB(&H61)&ChrB(&H6D)&ChrB(&H65)&ChrB(&H3D)&ChrB(&H22))
			y=InstrB(x+11,binFormData,ChrB(&H22))
			For z=1 to y-x-10
				strFileName=strFileName & Chr(AscB(MidB(binFormData,x+z+9,1)))
			Next
		End If

		Dim objStreamUp
		Set objStreamUp = Server.CreateObject("ADODB.Stream")

		With objStreamUp
			.Type = adTypeBinary
			.Mode = adModeReadWrite
			.Open
			.Position = 0
			.Write binFormData
			.Position = i
			Stream=.Read(j-i-1)
			.Close
		End With

		FileSize=LenB(Stream)

	End Function


	Function UpLoad_Stream()

		FileSize=LenB(Stream)

	End Function


	Public Function UpLoad(bolAutoName)

		If UploadType="Form" Then
			Call UpLoad_Form()
		ElseIf UploadType="Stream" Then
			Call UpLoad_Stream()
		End If

		If bolAutoName=True Then
		End If

		Dim objStreamFile
		Set objStreamFile = Server.CreateObject("ADODB.Stream")

		objStreamFile.Type = adTypeBinary
		objStreamFile.Mode = adModeReadWrite
		objStreamFile.Open
		objStreamFile.Write Stream
		objStreamFile.SaveToFile FileName,adSaveCreateOverWrite
		objStreamFile.Close


		UpLoad=True

	End Function


	Public Function Del()

		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")

		If fso.FileExists(FileName) Then
			fso.DeleteFile(FileName)
		End If

		Del=True
		
	End Function


End Class
'*********************************************************
%>
<!--以下是显示页面-->
<%
Action=Request.QueryString("act")
SelectedPlugin=Request.QueryString("Plugin")
SelectedPluginName=Request.QueryString("Pluginname")
If SelectedPluginName = "" Then SelectedPluginName = SelectedPlugin

If Action="" Then
Call GetBlogHint()

Response.Write "<div>"
%>
<form border="1" name="edit" id="edit" method="post" enctype="multipart/form-data" action="Xml_Upload.asp?act=FileUpload"><p>选择本地的 ZPI 插件安装包文件,TS 将从该文件导入插件并安装到 PLUGIN 目录下: </p><p><input type="file" id="edtFileLoad" name="edtFileLoad" size="25">  <input type="submit" class="button" value="提交" name="B1" /> <input class="button" type="reset" value="重置" name="B2" /> <input onclick="self.location.href='PluginList.asp'" type="button" class="button" value="返回插件管理" title="返回插件管理页" /></p>
<br />
</form>
<p><a href="help.asp#aboutzpi" title="什么是 ZPI 插件安装包文件?">[什么是 ZPI 插件安装包文件?]</a></p>
<%

Response.Write "</div>"
End If

Dim Install_Error
Install_Error=0

Dim Install_Pack,Install_Path
Install_Pack = BlogPath & "ZB_USERS/PLUGIN/Install.zpi"
Install_Path = BlogPath & "ZB_USERS/PLUGIN/"


'从本地上传
If Action="FileUpload" Then

	Response.Write "<p id=""loading"">正在导入插件, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<p class=""status-box"">正在上传 ZPI 插件安装包文件... <img id=""status"" align=""absmiddle"" src=""images/loading.gif"" /></p>"
	Response.Flush

	Dim objUpLoadFile
	Set objUpLoadFile=New TUpLoadFile2

	objUpLoadFile.FileName=Install_Pack
	objUpLoadFile.Del

	If objUpLoadFile.Upload(false)=False Then
		Response.Write "<p><font color=""red""> × ZPI 文件上传失败.</font></p>"
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('status').style.display = 'none';</script>"
		Install_Error=Install_Error+1
	Else
		Response.Write "<p><font color=""green""> √ ZPI 文件 ""PluginS/Install.ZPI"" 已被保存到您的空间内.</font></p>"
		Response.Write "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('status').style.display = 'none';</script>"
		Response.Flush
	End If

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"

	Set objUpLoadFile=Nothing

	Call Check_Install()

	Call Install_Plugin()

End If

If Action="continue" Then
	Call Install_Plugin()
End If

If Action="cancel" Then
	Call DeleteFile(Install_Pack)
	Response.Write "<p class=""status-box"">插件安装已取消. 如果您的浏览器没能自动跳转, 请 <a href=""xml_Upload.asp"">[点击这里]</a>.</p>"
	Response.Write "<script>setTimeout(""self.location.href='xml_Upload.asp'"",1000);</script>"
End If

'*********************************************************
Sub Check_Install()
On Error Resume Next

	If Install_Error<>0 Then
		Response.Write "<p class=""status-box""><font color=""red""> × 插件上传失败. "
		Response.Write "请 <a href=""Xml_upload.asp?confirm=yes&amp;act=cancel"" title=""取消安装并返回""><span>[取消安装并返回]</span></a></font></p>"
		Response.End
	End If

	Dim Alert
	Alert=False

	Response.Write "<p id=""loading2"">正在校验插件, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Set objXmlVerChk=New PluginSapper_CheckVersionViaXML
	objXmlVerChk.XmlDataWeb=(LoadFromFile(Install_Pack,"utf-8"))
	objXmlVerChk.XmlDataLocal=(LoadFromFile(BlogPath & "/ZB_USERS/PLUGIN/"& objXmlVerChk.Item_ID_Web &"/plugin.xml","utf-8"))

	If LCase(objXmlVerChk.Item_ID_Web)=LCase(objXmlVerChk.Item_ID_Local) And Install_Error=0 Then
		Response.Write "<p class=""status-box"">您已安装了这个插件 <b>("& objXmlVerChk.Item_Name_Local &")</b>, 是否用 ZPI 文件里的插件 <b>("& objXmlVerChk.Item_Name_Web &")</b> <b>完全覆盖</b>已安装的插件?<br/><br/>"

		Response.Write "您当前插件版本为: <b>"& objXmlVerChk.Item_Version_Local &"</b>. 发布日期为: <b>"& objXmlVerChk.Item_PubDate_Local &"</b>. 最后修改日期为: <b>"& objXmlVerChk.Item_Modified_Local &"</b>.<br/>"
		Response.Write "将要覆盖的版本为: <b>"& objXmlVerChk.Item_Version_Web &"</b>. 发布日期为: <b>"& objXmlVerChk.Item_PubDate_Web &"</b>. 最后修改日期为: <b>"& objXmlVerChk.Item_Modified_Web &"</b><br/><br/>"

		If objXmlVerChk.Item_Url_Web<>Empty Then
			Response.Write "<a href="""& objXmlVerChk.Item_Url_Web &""" target=""_blank"" title=""查看插件的发布页面"">点此查看插件的发布信息!</a><br/><br/>"
		End If

		Response.Write objXmlVerChk.OutputResults & "<br/><br/>"

		Response.Write "<a href=""Xml_upload.asp?confirm=yes&amp;act=continue"" title=""确认安装"">[确认覆盖]</a> 或 <a href=""Xml_upload.asp?confirm=yes&amp;act=cancel"" title=""取消安装"">[取消]</a></p>"
		Alert=True
	End If

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading2').style.display = 'none';}catch(e){};</script>"

	Set objXmlVerChk=Nothing

	If Alert=True Then Response.End

End Sub


Sub Install_Plugin()
On Error Resume Next

	Response.Write "<p id=""loading3"">正在解包插件, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Response.Write "<p class=""status-box"">ZPI 文件 ""PLUGIN/Install.zpi"" 正在解包安装...<p>"
	Response.Flush

	Dim objXmlFile
	Dim objNodeList
	Dim objFSO
	Dim objStream
	Dim i,j

	Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")
		objXmlFile.async = False
		objXmlFile.ValidateOnParse=False
		objXmlFile.load(Install_Pack)
		
		If objXmlFile.readyState<>4 Then
			Response.Write "<p><font color=""red""> × ZPI 文件未准备就绪, 无法解包.</font></p>"
			Install_Error=Install_Error+1
		Else
			If objXmlFile.parseError.errorCode <> 0 Then
				Response.Write "<p><font color=""red""> × ZPI 文件有错误, 无法解包.</font></p>"
				Install_Error=Install_Error+1
			Else

				Dim Pack_ver,Pack_Type,Pack_For,Pack_ID,Pack_Name
				Pack_Ver = objXmlFile.documentElement.SelectSingleNode("//root").getAttributeNode("version").value
				Pack_Type = objXmlFile.documentElement.selectSingleNode("//root").getAttributeNode("type").value
				Pack_For = objXmlFile.documentElement.selectSingleNode("//root").getAttributeNode("for").value
				Pack_ID = objXmlFile.documentElement.selectSingleNode("id").text
				Pack_Name = objXmlFile.documentElement.selectSingleNode("name").text

				If (CDbl(Pack_Ver) > CDbl(XML_Pack_Ver)) Then
					Response.Write "<p><font color=""red""> × ZPI 文件的 XML 版本为 "& Pack_Ver &", 而你的解包器版本为 "& XML_Pack_Ver &", 请升级您的 PluginSapper, 安装被中止.</font></p>"
					Install_Error=Install_Error+1
				ElseIf (LCase(Pack_Type) <> LCase(XML_Pack_Type)) Then
					Response.Write "<p><font color=""red""> × 不是 ZPI 文件, 而可能是 "& Pack_Type &", 安装被中止.</font></p>"
					Install_Error=Install_Error+1
				ElseIf (LCase(Pack_For) <> LCase(XML_Pack_Version)) Then
					Response.Write "<p><font color=""red""> × ZPI 文件版本不符合, 该版本可能是 "& Pack_For &", 安装被中止.</font></p>"
					Install_Error=Install_Error+1
				Else

					Response.Write "<blockquote><font color=""Teal"">"

					Set objNodeList = objXmlFile.documentElement.selectNodes("//folder/path")
					Set objFSO = CreateObject("Scripting.FileSystemObject")
						
						j=objNodeList.length-1
						For i=0 To j
							If objFSO.FolderExists(Install_Path & objNodeList(i).text)=False Then
								objFSO.CreateFolder(Install_Path & objNodeList(i).text)
							End If
							Response.Write "创建目录" & objNodeList(i).text & "<br/>"
							Response.Flush
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
									Response.Write "释放文件" & objNodeList(i).text & "<br/>"
									Response.Flush
									.Close
								End With
							Set objStream = Nothing
						Next
					Set objNodeList = Nothing
					Response.Write "</font></blockquote>"

				End If

			End If
		End If
			
		Set objXmlFile = Nothing

		If Err.Number<>0 Then Install_Error=Install_Error+1
		Err.Clear

		Response.Write "<p>"
		Install_Error = Install_Error + DeleteFile(Install_Pack)
		Response.Write "</p>"

		If Install_Error = 0 Then
			Response.Write "<p>"
			Install_Error = Install_Error + DeleteFile(BlogPath & "ZB_USERS/PLUGIN/" & Pack_ID & "/verchk.xml")
			Response.Write "</p>"
		End If

		Response.Flush


	If Install_Error = 0 Then
		Response.Write "<p class=""status-box""> √ 插件导入完成. 如果您的浏览器没能自动跳转, 请 <a href=""PluginDetail.asp?plugin=" & Server.URLEncode(Pack_ID) & "&amp;pluginname=" & Server.URLEncode(Pack_Name) & """>[点击这里]</a>.</p>"
		Response.Write "<script>setTimeout(""self.location.href='PluginDetail.asp?plugin=" & Server.URLEncode(Pack_ID) & "&pluginname=" & Server.URLEncode(Pack_Name) & "'"",3000);</script>"
	Else
		Response.Write "<p class=""status-box""><font color=""red""> × 插件导入失败. "
		Response.Write "请 <a href=""Xml_upload.asp?confirm=yes&amp;act=cancel"" title=""取消安装并返回""><span>[取消安装并返回]</span></a></font></p>"
	End If

	Response.Write "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading3').style.display = 'none';}catch(e){};</script>"

End Sub
'*********************************************************
%>
</div>
<script>


</script>
</body>
</html>
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>