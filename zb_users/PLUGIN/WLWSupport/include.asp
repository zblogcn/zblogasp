<%

'注册插件
Call RegisterPlugin("WLWSupport","ActivePlugin_WLWSupport")


'具体的接口挂接
Function ActivePlugin_WLWSupport() 

	'挂上接口
	'Action_Plugin_XMLRPC_Begin
	Call Add_Action_Plugin("Action_Plugin_XMLRPC_Begin","Call WLWSupport_Main()")


	'Action_Plugin_TArticleList_ExportByCache_Begin
	Call Add_Action_Plugin("Action_Plugin_TArticleList_Export_Begin","Call Add_Filter_Plugin(""Filter_Plugin_TArticleList_Build_Template"",""WLWSupport_EditDefault"")")

End Function


Function InstallPlugin_WLWSupport()

	On Error Resume Next

	BlogReBuild_Default
	'Call SetBlogHint(True,True,Empty)

	Err.Clear

End Function


Function UninstallPlugin_WLWSupport()

	On Error Resume Next

	'Call SetBlogHint(True,True,Empty)

	Err.Clear

End Function



Function WLWSupport_EditDefault(ByRef html)

	Dim s,t
	s="<link rel=""EditURI"" type=""application/rsd+xml"" href="""& ZC_BLOG_HOST &"zb_users/plugin/wlwsupport/rsd.asp"" />"
	t="<link rel=""wlwmanifest"" type=""application/wlwmanifest+xml"" href="""& ZC_BLOG_HOST &"zb_users/plugin/wlwsupport/wlwmanifest.asp"" />"

	html=Replace(html,"</head>",s & vbCrlf & t & vbCrlf &"</head>",1,-1,1)

End Function


Function WLWSupport_Main() 

	Response.ContentType = "text/xml"

	Dim strXmlCall
	Dim objXmlFile

	strXmlCall=Request.BinaryRead(Request.TotalBytes)
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
					If CheckUserAndRights(strUserName,strUserPassWord,"ArticleMng") Then Call WLWSupport_getRecentPosts(intNumberOfPosts)
				Case "metaWeblog.newPost":
					strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
					strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
					strPost=objRootNode.selectSingleNode("params/param[3]/value/struct").xml
					bolPublish=CBool(objRootNode.selectSingleNode("params/param[4]/value/boolean").text)
					If CheckUserAndRights(strUserName,strUserPassWord,"ArticleEdt") Then Call WLWSupport_newPost(strPost,bolPublish)
				Case "metaWeblog.editPost":
					intPostID=objRootNode.selectSingleNode("params/param[0]/value/string").text
					strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
					strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
					strPost=objRootNode.selectSingleNode("params/param[3]/value/struct").xml
					bolPublish=CBool(objRootNode.selectSingleNode("params/param[4]/value/boolean").text)
					If CheckUserAndRights(strUserName,strUserPassWord,"ArticleEdt") Then Call WLWSupport_editPost(intPostID,strPost,bolPublish)
				Case "metaWeblog.getPost":
					intPostID=objRootNode.selectSingleNode("params/param[0]/value/string").text
					strUserName=objRootNode.selectSingleNode("params/param[1]/value/string").text
					strUserPassWord=objRootNode.selectSingleNode("params/param[2]/value/string").text
					If CheckUserAndRights(strUserName,strUserPassWord,"ArticleMng") Then Call WLWSupport_getPost(intPostID)
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

	Response.End

End Function





'*********************************************************
' 目的：    
'*********************************************************
Function WLWSupport_getRecentPosts(numberOfPosts)

	Dim strXML
	Dim strPost
	Dim strRecentPosts

	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value><array><data>$%#1#%$</data></array></value></param></params></methodResponse>"

	strPost="<value><struct><member><name>title</name><value><string>$%#1#%$</string></value></member><member><name>description</name><value><string>$%#2#%$</string></value></member><member><name>dateCreated</name><value><dateTime.iso8601>$%#3#%$</dateTime.iso8601></value></member><member><name>categories</name><value><array><data><value><string>$%#4#%$</string></value></data></array></value></member><member><name>postid</name><value><string>$%#5#%$</string></value></member><member><name>userid</name><value><string>$%#6#%$</string></value></member><member><name>link</name><value><string>$%#7#%$</string></value></member><member><name>mt_excerpt</name><value><string>$%#8#%$</string></value></member><member><name>mt_text_more</name><value><string></string></value></member><member><name>mt_allow_comments</name><value><int></int></value></member><member><name>mt_allow_pings</name><value><int></int></value></member><member><name>mt_basename</name><value><string>$%#9#%$</string></value></member><member><name>mt_keywords</name><value><string>$%#10#%$</string></value></member></struct></value>"

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

				s=Replace(s,"$%#8#%$",TransferHTML(objArticle.Intro,"[html-format]"))
				s=Replace(s,"$%#9#%$",TransferHTML(objArticle.Alias,"[html-format]"))
				s=Replace(s,"$%#10#%$",TransferHTML(objArticle.TagToName,"[html-format]"))

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
Function WLWSupport_newPost(structPost,bolPublish)

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

	If objXmlFile.documentElement.SelectNodes("member[name=""categories""]/value/array/data/value[1]/string").Count > 0 Then
		strCate=objXmlFile.documentElement.selectSingleNode("member[name=""categories""]/value/array/data/value[0]/string").text
	End If

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

	objArticle.Alias=objXmlFile.documentElement.selectSingleNode("member[name=""mt_basename""]/value/string").text

	objArticle.Tag=ParseTag(objXmlFile.documentElement.selectSingleNode("member[name=""mt_keywords""]/value/string").text)

	If objXmlFile.documentElement.SelectNodes("member[name=""dateCreated""]/value/dateTime.iso8601").count>0 Then
		Dim y,m,d,t,dt
		y=Mid(objXmlFile.documentElement.selectSingleNode("member[name=""dateCreated""]/value/dateTime.iso8601").text,1,4)
		m=Mid(objXmlFile.documentElement.selectSingleNode("member[name=""dateCreated""]/value/dateTime.iso8601").text,5,2)
		d=Mid(objXmlFile.documentElement.selectSingleNode("member[name=""dateCreated""]/value/dateTime.iso8601").text,7,2)
		t=Mid(objXmlFile.documentElement.selectSingleNode("member[name=""dateCreated""]/value/dateTime.iso8601").text,10,8)
		dt=y & "-" & m & "-" & d & " " & t
		If IsDate(dt)=True Then
			objArticle.PostTime=dt
		End If
	End If

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

	If Trim(objXmlFile.documentElement.selectSingleNode("member[name=""mt_excerpt""]/value/string").text)<>"" Then
		objArticle.Intro=objXmlFile.documentElement.selectSingleNode("member[name=""mt_excerpt""]/value/string").text
	End If

	If objArticle.Post=True Then
		Call BuildArticle(objArticle.ID,true,true)

		Call MakeBlogReBuild()
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
Function WLWSupport_editPost(intPostID,structPost,bolPublish)

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

	If objXmlFile.documentElement.SelectNodes("member[name=""categories""]/value/array/data/value[1]/string").Count > 0 Then
		strCate=objXmlFile.documentElement.selectSingleNode("member[name=""categories""]/value/array/data/value[0]/string").text
	End If

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

	objArticle.Alias=objXmlFile.documentElement.selectSingleNode("member[name=""mt_basename""]/value/string").text

	objArticle.Tag=ParseTag(objXmlFile.documentElement.selectSingleNode("member[name=""mt_keywords""]/value/string").text)

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

	If objXmlFile.documentElement.SelectNodes("member[name=""dateCreated""]/value/dateTime.iso8601").count>0 Then
		Dim y,m,d,t,dt
		y=Mid(objXmlFile.documentElement.selectSingleNode("member[name=""dateCreated""]/value/dateTime.iso8601").text,1,4)
		m=Mid(objXmlFile.documentElement.selectSingleNode("member[name=""dateCreated""]/value/dateTime.iso8601").text,5,2)
		d=Mid(objXmlFile.documentElement.selectSingleNode("member[name=""dateCreated""]/value/dateTime.iso8601").text,7,2)
		t=Mid(objXmlFile.documentElement.selectSingleNode("member[name=""dateCreated""]/value/dateTime.iso8601").text,10,8)
		dt=y & "-" & m & "-" & d & " " & t
		If IsDate(dt)=True Then
			objArticle.PostTime=dt
		End If
	End If

	If Trim(objXmlFile.documentElement.selectSingleNode("member[name=""mt_excerpt""]/value/string").text)<>"" Then
		objArticle.Intro=objXmlFile.documentElement.selectSingleNode("member[name=""mt_excerpt""]/value/string").text
	End If

	If objArticle.Post=True Then
		Call BuildArticle(objArticle.ID,true,true)

		Call MakeBlogReBuild()
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
Function WLWSupport_getPost(intPostID)

	Dim strXML
	Dim strPost
	Dim strRecentPosts
	Dim s

	strXML="<?xml version=""1.0"" encoding=""UTF-8""?><methodResponse><params><param><value>$%#1#%$</value></param></params></methodResponse>"

	strPost="<struct><member><name>title</name><value><string>$%#1#%$</string></value></member><member><name>description</name><value><string>$%#2#%$</string></value></member><member><name>dateCreated</name><value><dateTime.iso8601>$%#3#%$</dateTime.iso8601></value></member><member><name>categories</name><value><array><data><value><string>$%#4#%$</string></value></data></array></value></member><member><name>postid</name><value><string>$%#5#%$</string></value></member><member><name>userid</name><value><string>$%#6#%$</string></value></member><member><name>link</name><value><string>$%#7#%$</string></value></member><member><name>mt_excerpt</name><value><string>$%#8#%$</string></value></member><member><name>mt_text_more</name><value><string></string></value></member><member><name>mt_allow_comments</name><value><int></int></value></member><member><name>mt_allow_pings</name><value><int></int></value></member><member><name>mt_basename</name><value><string>$%#9#%$</string></value></member><member><name>mt_keywords</name><value><string>$%#10#%$</string></value></member></struct>"


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

	s=Replace(s,"$%#8#%$",TransferHTML(objArticle.Intro,"[html-format]"))
	s=Replace(s,"$%#9#%$",TransferHTML(objArticle.Alias,"[html-format]"))
	s=Replace(s,"$%#10#%$",TransferHTML(objArticle.TagToName,"[html-format]"))

	strRecentPosts=strRecentPosts & s

	strXML=Replace(strXML,"$%#1#%$",strRecentPosts)

	Response.Write strXML

End Function
'*********************************************************

%>