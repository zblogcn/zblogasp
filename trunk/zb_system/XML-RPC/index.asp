<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    XML-RPC/index.asp
'// 开始时间:    2005.09.30
'// 最后修改:    
'// 备    注:    XML-RPC主文件
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_function_md5.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_event.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../plugin/p_config.asp" -->
<!-- #include file="../function/rss_lib.asp" -->
<!-- #include file="../function/atom_lib.asp" -->
<%
'/////////////////////////////////////////////////////////////////////////////////////////
'*********************************************************
' 目的：    
'*********************************************************
Function ParseDateForRFC3339(dtmDate)

	Dim dtmDay, dtmWeekDay, dtmMonth, dtmYear
	Dim dtmHours, dtmMinutes, dtmSeconds

	Dim strTimeZone

	dtmYear = Year(dtmDate)
	dtmMonth = Right("00" & Month(dtmDate),2)
	dtmDay = Right("00" & Day(dtmDate),2)

	dtmHours = Right("00" & Hour(dtmDate),2)
	dtmMinutes = Right("00" & Minute(dtmDate),2)
	dtmSeconds = Right("00" & Second(dtmDate),2)

	strTimeZone=Left(ZC_TIME_ZONE,3) & ":" & Right(ZC_TIME_ZONE,2)

	ParseDateForRFC3339 = dtmYear & "-" & dtmMonth & "-" & dtmDay & "T" & dtmHours & ":" & dtmMinutes & ":" & dtmSeconds & strTimeZone

End Function 
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function CheckUserAndRights(userName,userPassWord,strAction)

	Set BlogUser=Nothing
	Set BlogUser=New TUser

	BlogUser.LoginType="Self"
	BlogUser.Name=userName
	BlogUser.PassWord=MD5(userPassWord)
	If BlogUser.Verify() Then
		If Not CheckRights(strAction) Then Call RespondError(6,ZVA_ErrorMsg(6))
		CheckUserAndRights=True
	Else
		Call RespondError(7,ZVA_ErrorMsg(7))
	End If

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function this_getUsersBlogs()

	Dim strXML
	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><array><data><value><struct><member><name>url</name><value><string>$%#1#%$</string></value></member><member><name>blogid</name><value><string>$%#2#%$</string></value></member><member><name>blogName</name><value><string>$%#3#%$</string></value></member></struct></value></data></array></value></param></params></methodResponse>"

	strXML=Replace(strXML,"$%#1#%$",TransferHTML(ZC_BLOG_HOST,"[html-format]"))
	strXML=Replace(strXML,"$%#2#%$",TransferHTML(ZC_BLOG_CLSID,"[html-format]"))
	strXML=Replace(strXML,"$%#3#%$",TransferHTML(ZC_BLOG_NAME,"[html-format]"))

	Response.Write strXML

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function this_getCategories()

	Dim strXML
	Dim strCategoryInfo

	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><array><data>$%#1#%$</data></array></value></param></params></methodResponse>"

	strCategoryInfo="<value><struct><member><name>description</name><value><string>$%#1#%$</string></value></member><member><name>httpUrl</name><value><string>$%#2#%$</string></value></member><member><name>rssUrl</name><value><string>$%#3#%$</string></value></member><member><name>title</name><value><string>$%#4#%$</string></value></member><member><name>categoryid</name><value><string>$%#5#%$</string></value></member></struct></value>"

	Dim Cate
	Dim s
	Dim strCategories
	For Each Cate in Categorys
		If IsObject(Cate) Then
			s=strCategoryInfo
			s=Replace(s,"$%#1#%$",TransferHTML(Cate.Name,"[html-format]"))
			s=Replace(s,"$%#2#%$",TransferHTML(Cate.Url,"[html-format]"))
			s=Replace(s,"$%#3#%$",TransferHTML(Cate.Url,"[html-format]"))
			s=Replace(s,"$%#4#%$",TransferHTML(Cate.Name,"[html-format]"))
			s=Replace(s,"$%#5#%$",TransferHTML(Cate.ID,"[html-format]"))
			strCategories=strCategories & s
		End If
	Next

	strXML=Replace(strXML,"$%#1#%$",strCategories)

	Response.Write strXML

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function this_getRecentPosts(numberOfPosts)

	Dim strXML
	Dim strPost
	Dim strRecentPosts

	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><array><data>$%#1#%$</data></array></value></param></params></methodResponse>"

	strPost="<value><struct><member><name>title</name><value><string>$%#1#%$</string></value></member><member><name>description</name><value><string>$%#2#%$</string></value></member><member><name>dateCreated</name><value><dateTime.iso8601>$%#3#%$</dateTime.iso8601></value></member><member><name>categories</name><value><array><data><value><string>$%#4#%$</string></value></data></array></value></member><member><name>postid</name><value><string>$%#5#%$</string></value></member><member><name>userid</name><value><string>$%#6#%$</string></value></member><member><name>link</name><value><string>$%#7#%$</string></value></member></struct></value>"

	Dim s
	Dim i
	Dim objRS
	Dim strSQL
	Dim strPage
	Dim objArticle

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""

	If CheckRights("Root")=False Then strSQL="WHERE [log_AuthorID] = " & BlogUser.ID

	objRS.Open("SELECT [log_ID] FROM [blog_Article] "& strSQL &" ORDER BY [log_PostTime] DESC")
	objRS.PageSize=numberOfPosts
	If objRS.PageCount>0 Then objRS.AbsolutePage = 1

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize

			Set objArticle=New TArticle

			If objArticle.LoadInfoByID(objRS("log_ID")) Then
				s=strPost
				s=Replace(s,"$%#1#%$",TransferHTML(objArticle.Title,"[html-format]"))
				s=Replace(s,"$%#2#%$",TransferHTML(objArticle.Content,"[html-format]"))
				s=Replace(s,"$%#3#%$",TransferHTML(ParseDateForRFC3339(objArticle.PostTime),"[html-format]"))
				s=Replace(s,"$%#4#%$",TransferHTML(Categorys(objArticle.CateID).Name,"[html-format]"))
				s=Replace(s,"$%#5#%$",TransferHTML(objArticle.ID,"[html-format]"))
				s=Replace(s,"$%#6#%$",TransferHTML(objArticle.AuthorID,"[html-format]"))
				s=Replace(s,"$%#7#%$",TransferHTML(objArticle.Url,"[html-format]"))

				strRecentPosts=strRecentPosts & s
			End If

			objRS.MoveNext
			If objRS.eof Then Exit For

			Set objArticle=Nothing

		Next

	End If

	strXML=Replace(strXML,"$%#1#%$",strRecentPosts)

	Response.Write strXML

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function this_newPost(structPost,bolPublish)

	On Error Resume Next

	Dim objXmlFile
	Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")

	objXmlFile.loadXML(structPost)

	Dim strXML

	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><string>$%#1#%$</string></value></param></params></methodResponse>"


	Dim objArticle
	Set objArticle=New TArticle
	objArticle.ID=0
	objArticle.AuthorID=BlogUser.ID
	If bolPublish=True Then
		objArticle.Level=4
	Else
		objArticle.Level=1
	End If
	objArticle.PostTime=Now()
	objArticle.Title=objXmlFile.documentElement.selectSingleNode("member[name=""title""]/value/string").text
	objArticle.Tag=""
	objArticle.Alias=""


	Dim strCate
	strCate=objXmlFile.documentElement.selectSingleNode("member[name=""categories""]/value/array/data/value[0]/string").text

	Dim Cate
	Dim i
	For i=UBound(Categorys) To LBound(Categorys) Step -1
		If IsObject(Categorys(i)) Then
			objArticle.CateID=Categorys(i).ID
			If strCate=Categorys(i).Name Then
				objArticle.CateID=Categorys(i).ID
				Exit For
			End If
		End If
	Next

	objArticle.Content=objXmlFile.documentElement.selectSingleNode("member[name=""description""]/value/string").text

	Dim objRegExp
	Dim s
	s=objArticle.Content
	Set objRegExp=New RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True
	objRegExp.Pattern="<[^>]*>"
	s=objRegExp.Replace(s,"")
	s=Left(s,ZC_TB_EXCERPT_MAX) & "..."
	objArticle.Intro=objArticle.Content

	If objArticle.Post=True Then
		Call BuildArticle(objArticle.ID,true,true)

		Call MakeBlogReBuild_Core()
		Response.Clear

		strXML=Replace(strXML,"$%#1#%$",objArticle.ID)
		Response.Write strXML
	Else
		Call RespondError(11,ZVA_ErrorMsg(11))
	End If

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function this_editPost(intPostID,structPost,bolPublish)

	On Error Resume Next

	Dim objXmlFile
	Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")

	objXmlFile.loadXML(structPost)

	Dim strXML

	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><boolean>$%#1#%$</boolean></value></param></params></methodResponse>"

	Dim objArticle
	Set objArticle=New TArticle

	If objArticle.LoadInfoByID(intPostID) Then
		If Not((objArticle.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Call RespondError(6,ZVA_ErrorMsg(6))
	Else
		Call RespondError(9,ZVA_ErrorMsg(9))
	End If

	objArticle.Title=objXmlFile.documentElement.selectSingleNode("member[name=""title""]/value/string").text

	If bolPublish=True Then
		objArticle.Level=4
	Else
		objArticle.Level=1
	End If

	Dim strCate
	strCate=objXmlFile.documentElement.selectSingleNode("member[name=""categories""]/value/array/data/value[0]/string").text
	If strCate<>"" Then
		Dim Cate
		Dim i
		For i=UBound(Categorys) To LBound(Categorys) Step -1
			If IsObject(Categorys(i)) Then
				objArticle.CateID=Categorys(i).ID
				If strCate=Categorys(i).Name Then
					objArticle.CateID=Categorys(i).ID
					Exit For
				End If
			End If
		Next
	End If
	objArticle.Content=objXmlFile.documentElement.selectSingleNode("member[name=""description""]/value/string").text

	Dim objRegExp
	Dim s
	s=objArticle.Content
	Set objRegExp=New RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True
	objRegExp.Pattern="<[^>]*>"
	s=objRegExp.Replace(s,"")
	s=Left(s,ZC_TB_EXCERPT_MAX) & "..."
	objArticle.Intro=objArticle.Content

	If objArticle.Post=True Then
		Call BuildArticle(objArticle.ID,true,true)

		Call MakeBlogReBuild_Core()
		Response.Clear

		strXML=Replace(strXML,"$%#1#%$",1)
		Response.Write strXML
	Else
		Call RespondError(11,ZVA_ErrorMsg(11))
	End If

	Err.Clear

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function this_getPost(intPostID)

	Dim strXML
	Dim strPost
	Dim strRecentPosts
	Dim s

	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value>$%#1#%$</value></param></params></methodResponse>"

	strPost="<struct><member><name>title</name><value><string>$%#1#%$</string></value></member><member><name>description</name><value><string>$%#2#%$</string></value></member><member><name>dateCreated</name><value><dateTime.iso8601>$%#3#%$</dateTime.iso8601></value></member><member><name>categories</name><value><array><data><value><string>$%#4#%$</string></value></data></array></value></member><member><name>postid</name><value><string>$%#5#%$</string></value></member><member><name>userid</name><value><string>$%#6#%$</string></value></member><member><name>link</name><value><string>$%#7#%$</string></value></member></struct>"


	Dim objArticle
	Set objArticle=New TArticle

	If objArticle.LoadInfoByID(intPostID) Then
		If Not((objArticle.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Call RespondError(6,ZVA_ErrorMsg(6))
	Else
		Call RespondError(9,ZVA_ErrorMsg(9))
	End If

	s=strPost
	s=Replace(s,"$%#1#%$",TransferHTML(objArticle.Title,"[html-japan][html-format]"))
	s=Replace(s,"$%#2#%$",TransferHTML(objArticle.Content,"[html-japan][html-format]"))
	s=Replace(s,"$%#3#%$",TransferHTML(ParseDateForRFC3339(objArticle.PostTime),"[html-format]"))
	s=Replace(s,"$%#4#%$",TransferHTML(Categorys(objArticle.CateID).Name,"[html-format]"))
	s=Replace(s,"$%#5#%$",TransferHTML(objArticle.ID,"[html-format]"))
	s=Replace(s,"$%#6#%$",TransferHTML(objArticle.AuthorID,"[html-format]"))
	s=Replace(s,"$%#7#%$",TransferHTML(objArticle.Url,"[html-format]"))

	strRecentPosts=strRecentPosts & s

	strXML=Replace(strXML,"$%#1#%$",strRecentPosts)

	Response.Write strXML

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function this_newMediaObject(strFileName,strFileType,strFileBits)

	Dim objXmlFile
	Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")

	Dim strXML
	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><struct><member><name>url</name><value><string>$%#1#%$</string></value></member></struct></value></param></params></methodResponse>"

	Dim objUpLoadFile
	Set objUpLoadFile=New TUpLoadFile
	objUpLoadFile.AuthorID=BlogUser.ID
	objUpLoadFile.FileName=strFileName
	objUpLoadFile.UploadType="Stream"

	If Not CheckRegExp(LCase(strFileName),"\.("& ZC_UPLOAD_FILETYPE &")$") Then Call RespondError(26,ZVA_ErrorMsg(26))

	Dim xmlnode
	Set xmlnode = objXmlFile.createElement("file")
	xmlnode.datatype = "bin.base64"
	xmlnode.text = strFileBits

	Dim objStreamUp
	Set objStreamUp = Server.CreateObject("ADODB.Stream")

	With objStreamUp
		.Type = adTypeBinary
		.Mode = adModeReadWrite
		.Open
		.Position = 0
		.Write xmlnode.nodeTypedvalue

		If .Size>ZC_UPLOAD_FILESIZE Then Call RespondError(27,ZVA_ErrorMsg(27))

		Dim objRS
		strFileName=FilterSQL(strFileName)
		If Not objConn.Execute("SELECT * FROM [blog_UpLoad] WHERE [ul_FileName] = '" & strFileName & "'").EOF Then Call RespondError(28,ZVA_ErrorMsg(28))

		.Position = 0
		objUpLoadFile.Stream=.Read
		.Close
	End With

	If objUpLoadFile.UpLoad(False) Then

		strXML=Replace(strXML,"$%#1#%$",TransferHTML(objUpLoadFile.FullUrlPathName,"[html-format]"))

		Response.Write strXML

	End If

End Function
'*********************************************************




'*********************************************************
' 目的：    
'*********************************************************
Function this_deletePost(intPostID)

	Dim strXML

	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><boolean>$%#1#%$</boolean></value></param></params></methodResponse>"

	Dim objArticle
	Set objArticle=New TArticle

	If objArticle.LoadInfoByID(intPostID) Then
		If Not((objArticle.AuthorID=BlogUser.ID) Or (CheckRights("Root")=True)) Then Call RespondError(6,ZVA_ErrorMsg(6))
	Else
		Call RespondError(9,ZVA_ErrorMsg(9))
	End If

	If objArticle.Del Then

		Call MakeBlogReBuild_Core()
		Response.Clear

		strXML=Replace(strXML,"$%#1#%$",1)
		Response.Write strXML
	Else
		Call RespondError(11,ZVA_ErrorMsg(11))
	End If

End Function
'*********************************************************




'/////////////////////////////////////////////////////////////////////////////////////////
Call System_Initialize()

Dim strXmlCall
Dim objXmlFile

'plugin node
For Each sAction_Plugin_XMLRPC_Begin in Action_Plugin_XMLRPC_Begin
	If Not IsEmpty(sAction_Plugin_XMLRPC_Begin) Then Call Execute(sAction_Plugin_XMLRPC_Begin)
Next


Response.ContentType = "text/xml"

If strXmlCall="" Then
	strXmlCall=Request.BinaryRead(Request.TotalBytes)
End If

Set objXmlFile = Server.CreateObject("Microsoft.XMLDOM")

objXmlFile.load(strXmlCall)

If objXmlFile.readyState=4 Then
	If objXmlFile.parseError.errorCode <> 0 Then
		Call RespondError(0,ZVA_ErrorMsg(0))
	Else

		Dim objRootNode
		Set objRootNode=objXmlFile.documentElement

		Dim strAction
		strAction=objRootNode.selectSingleNode("methodName").text

		Dim strUserName
		Dim strUserPassWord
		Dim intNumberOfPosts
		Dim strPost
		Dim intPostID
		Dim strFileName
		Dim strFileType
		Dim strFileBits
		Dim bolPublish

		Select Case strAction
			Case "blogger.getUsersBlogs":
				strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
				If CheckUserAndRights(strUserName,strUserPassWord,"admin") Then Call this_getUsersBlogs()
			Case "metaWeblog.getCategories":
				strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
				If CheckUserAndRights(strUserName,strUserPassWord,"admin") Then Call this_getCategories()
			Case "metaWeblog.getRecentPosts":
				strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
				intNumberOfPosts=objRootNode.selectSingleNode("params/param[3]/value/int").text
				If CheckUserAndRights(strUserName,strUserPassWord,"ArticleMng") Then Call this_getRecentPosts(intNumberOfPosts)
			Case "metaWeblog.newPost":
				strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
				strPost=objRootNode.selectSingleNode("params/param[3]/value/struct").xml
				bolPublish=CBool(objRootNode.selectSingleNode("params/param[4]/value/boolean").text)
				If CheckUserAndRights(strUserName,strUserPassWord,"ArticleEdt") Then Call this_newPost(strPost,bolPublish)
			Case "metaWeblog.editPost":
				intPostID=objRootNode.selectSingleNode("params/param[0]/value/string").text
				strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
				strPost=objRootNode.selectSingleNode("params/param[3]/value/struct").xml
				bolPublish=CBool(objRootNode.selectSingleNode("params/param[4]/value/boolean").text)
				If CheckUserAndRights(strUserName,strUserPassWord,"ArticleEdt") Then Call this_editPost(intPostID,strPost,bolPublish)
			Case "metaWeblog.getPost":
				intPostID=objRootNode.selectSingleNode("params/param[0]/value/string").text
				strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
				If CheckUserAndRights(strUserName,strUserPassWord,"ArticleMng") Then Call this_getPost(intPostID)
			Case "metaWeblog.newMediaObject":
				strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
				strFileName=objRootNode.selectSingleNode("params/param[3]/value/struct/member[name=""name""]/value/string").text
				strFileType=objRootNode.selectSingleNode("params/param[3]/value/struct/member[name=""type""]/value/string").text
				strFileBits=objRootNode.selectSingleNode("params/param[3]/value/struct/member[name=""bits""]/value/base64").text
				If CheckUserAndRights(strUserName,strUserPassWord,"FileUpload") Then Call this_newMediaObject(strFileName,strFileType,strFileBits)
			Case "blogger.deletePost":
				intPostID=objRootNode.selectSingleNode("params/param[1]/value/string").text
				strUserName=objRootNode.selectSingleNode("params/param[2]/value/string").text
				strUserPassWord=objRootNode.selectSingleNode("params/param[3]/value/string").text
				If CheckUserAndRights(strUserName,strUserPassWord,"ArticleDel") Then Call this_deletePost(intPostID)
			Case Else
				Call RespondError(1,ZVA_ErrorMsg(1))
		End Select 

	End If
End If

Call ClearGlobeCache
Call LoadGlobeCache

'plugin node
For Each sAction_Plugin_XMLRPC_End in Action_Plugin_XMLRPC_End
	If Not IsEmpty(sAction_Plugin_XMLRPC_End) Then Call Execute(sAction_Plugin_XMLRPC_End)
Next

Call System_Terminate()

If Err.Number<>0 then
	Call RespondError(0,ZVA_ErrorMsg(0))
End If
%>