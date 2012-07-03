<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.8 Pre Terminator 及以上版本, 其它版本的Z-blog未知
'// 插件制作:    haphic(http://haphic.com/)
'// 备    注:    主题管理插件
'// 最后修改：   2008-6-28
'// 最后版本:    1.2
'///////////////////////////////////////////////////////////////////////////////
Dim ThemeID,ThemeName,ThemeURL,ThemeNote,ThemePubDate
Dim ThemeAdapted,ThemeVersion,ThemeModified
Dim ThemeScreenShot,ThemeDescription

Dim ThemePlugin_Name,ThemePlugin_Note,ThemePlugin_Type
Dim ThemePlugin_Path,ThemePlugin_Include,ThemePlugin_Level

Dim ThemeSource_Name,ThemeSource_Url,ThemeSource_Email
Dim ThemeAuthor_Name,ThemeAuthor_Url,ThemeAuthor_Email

Dim Action,SelectedTheme,SelectedThemeName
Dim objXmlVerChk,NewVersionExists

Const DownLoad_URL="http://download.rainbowsoft.org/themes/ts.asp"
Const Resource_URL="http://download.rainbowsoft.org/Themes/"    '注意. Include 文件里还有一同名变量要修改
Const Update_URL="http://download.rainbowsoft.org/Plugin/dlcs/download.asp?theme="

Const XML_Pack_Ver="1.0"
Const XML_Pack_Type="Theme"
Const XML_Pack_Version="Z-Blog_1_8"

'定义超时时间
Const SiteResolve = 5    'UNISON_SiteResolve(Msxml2.ServerXMLHTTP有效)域名分析超时(秒)推荐为"5"	'提示 1秒=1000毫秒
Const SiteConnect = 5    'UNISON_SiteConnect(Msxml2.ServerXMLHTTP有效)连接站点超时(秒)推荐为"5"
Const SiteSend = 4    'UNISON_SiteSend(Msxml2.ServerXMLHTTP有效)发送数据时间超时(秒)推荐为"4"
Const SiteReceive = 10    'UNISON_SiteReceive(Msxml2.ServerXMLHTTP有效)等待反馈时间超时(秒)推荐为"10"

'***************************************************************************************



'***************************************************************************************
' 目的：    页面上部导航 
'***************************************************************************************
Sub SapperMenu(strCata)
	Dim Cata_1,Cata_2,Cata_3,Cata_4,Cata_5,Cata_8,Cata_9
	Cata_1="m-left":Cata_2="m-left":Cata_3="m-left":Cata_4="m-left":Cata_5="m-left":Cata_8="m-right":Cata_9="m-right"
	If strCata="1" Then Cata_1=Cata_1 & " m-now"
	If strCata="2" Then Cata_2=Cata_2 & " m-now"
	If strCata="3" Then Cata_3=Cata_3 & " m-now"
	If strCata="4" Then Cata_4=Cata_4 & " m-now"
	If strCata="5" Then Cata_5=Cata_5 & " m-now"
	If strCata="8" Then Cata_8=Cata_8 & " m-now"
	Response.Write "<div class=""SubMenu"">"


	Response.Write "<span class="""& Cata_1 &"""><a href="""&ZC_BLOG_HOST&"PLUGIN/ThemeSapper/Xml_List.asp"" title=""从服务器安装主题"">获取更多主题</a></span>"
	Response.Write "<span class="""& Cata_2 &"""><a href="""&ZC_BLOG_HOST&"PLUGIN/ThemeSapper/ThemeList.asp"" title=""主题管理页面"">主题管理扩展面板</a></span>"
	Response.Write "<span class="""& Cata_3 &"""><a href="""&ZC_BLOG_HOST&"PLUGIN/ThemeSapper/Xml_Upload.asp"" title=""从本地导入ZTI文件并安装主题"">从本地导入ZTI文件</a></span>"
	Response.Write "<span class="""& Cata_4 &"""><a href="""&ZC_BLOG_HOST&"PLUGIN/ThemeSapper/Xml_Restor.asp"" title=""管理主机上的ZTI文件"">管理主机上的ZTI文件</a></span>"
	Response.Write "<span class="""& Cata_5 &"""><a href="""&ZC_BLOG_HOST&"PLUGIN/ThemeSapper/Xml_ChkVer.asp"" title=""查看已安装主题的可用更新"">查看主题的可用更新</a></span>"

	Response.Write "<span class="""& Cata_9 &"""><a href="""&ZC_BLOG_HOST&"cmd.asp?act=ThemesMng"" title=""退出到插件管理页面"">退出 ThemeSapper</a></span>"
	Response.Write "<span class="""& Cata_8 &"""><a href="""&ZC_BLOG_HOST&"PLUGIN/ThemeSapper/help.asp"" title=""帮助文件"">帮助说明</a></span>"

	Response.Write "</div>"
end Sub
'***************************************************************************************




'*********************************************************
' 目的：    取得文件扩展名
'*********************************************************
Function GetFileExt(sFileName)
	GetFileExt = LCase(Mid(sFileName,InStrRev (sFileName, ".")+1))
End Function
'*********************************************************
' 目的：    检查某目录下的某文件是否存在
'*********************************************************
Function FileExists(fileName)
On Error Resume Next
Dim objFSO
FileExists = False
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(fileName) Then
	FileExists = True
End If
Set objFSO = Nothing
Err.Clear
End Function
'*********************************************************
' 目的：    复制文件
'*********************************************************
Function CopyFile(SFile,DFile)
On Error Resume Next
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
	fso.CopyFile SFile, DFile
If Err.Number = 53 Then
	CopyFile = 53
	Response.Write "<font color=""red""> × 安装文件 """& Replace(SFile,BlogPath,"") &"""不存在!</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Elseif Err.Number = 70 Then
	CopyFile = 70
	Response.Write "<font color=""red""> × 目标文件 """& Replace(DFile,BlogPath,"") &"""已存在且属性为只读!</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Elseif Err.Number <> 0 Then
	Response.Write "<font color=""red""> × 未知错误，错误编码：" & Err.Number & "</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Else
	Response.Write "<font color=""green""> √ 文件 """& Replace(DFile,BlogPath,"") &""" 创建成功.</font>"
	CopyFile = 0
End If
Set fso=Nothing
End Function
'*********************************************************
' 目的：    删除文件
'*********************************************************
Function DeleteFile(FileName)
On Error Resume Next
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile(FileName)
If Err.Number = 53 Then
	DeleteFile = 0
	Response.Write "<font color=""green""> √ 文件 """& Replace(FileName,BlogPath,"") &"""不存在!</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Elseif Err.Number = 70 Then
	DeleteFile = 70
	Response.Write "<font color=""red""> × 文件 """& Replace(FileName,BlogPath,"") &"""为只读, 无法删除!</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Elseif Err.Number <> 0 Then
	DeleteFile = Err.Number
	Response.Write "<font color=""red""> × 未知错误，错误编码：" & Err.Number & "</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Else
	Response.Write "<font color=""green""> √ 文件 """& Replace(FileName,BlogPath,"") &"""删除成功.</font>"
	DeleteFile = 0
End If
Set fso = Nothing
End Function
'*********************************************************
' 目的：    删除文件夹
'*********************************************************
Function DeleteFolder(FolderName)
on Error Resume Next
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
	fso.DeleteFolder(FolderName)
If Err.Number = 76 Then
	DeleteFolder = 0
	Response.Write "<font color=""green""> √ 文件夹 """& Replace(FolderName,BlogPath,"") &"""不存在!</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Elseif Err.Number = 70 Then
	DeleteFolder = 70
	Response.Write "<font color=""red""> × 文件夹 """& Replace(FolderName,BlogPath,"") &"""无法操作!</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Elseif Err.Number <> 0 Then
	DeleteFolder = Err.Number
	Response.Write "<font color=""red""> × 未知错误，错误编码：" & Err.Number & "</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Else
	Response.Write "<font color=""green""> √ 文件夹 """& Replace(FolderName,BlogPath,"") &"""删除成功.</font>"
	DeleteFolder = 0
End If
Set fso = Nothing
End Function

'*********************************************************
' 目的：    取得目标网页的html代码
'*********************************************************
Function getHTTPPage(url)
On Error Resume Next
Dim Http,ServerConn
Dim j
For j=0 To 2
	Set Http=server.createobject("Msxml2.ServerXMLHTTP")
	Http.setTimeouts SiteResolve*1000,SiteConnect*1000,SiteSend*1000,SiteReceive*1000
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
	getHTTPPage = False
	Exit Function
End If
If http.Status=200 Then
	'getHTTPPage=Http.ResponseText
	getHTTPPage=bytesToBSTR(Http.ResponseBody,"utf-8")
Else
	getHTTPPage = False
End If
Set http=Nothing
End Function
'*********************************************************
' 目的：    将目标网页转换为某种编码
'*********************************************************
Function BytesToBstr(strPageContent,strPageCharset)
On Error Resume Next
Dim objstream
Set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write strPageContent
objstream.Position = 0
objstream.Type = 2
objstream.CharSet = strPageCharset
BytesToBstr = objstream.ReadText
objstream.Close
Set objstream = Nothing
Err.Clear
End Function
'*********************************************************
%>