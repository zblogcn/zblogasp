<Script Language="VBScript" RunAt="Server">
'///////////////////////////////////////////////////////////////////////////////
'//              RainbowSoft ATOM Export
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    atom_lib.asp
'// 开始时间:    2005.07.27
'// 最后修改:    
'// 备    注:    ATOM输出
'///////////////////////////////////////////////////////////////////////////////




'*********************************************************
' 目的：    定义TAtom10Export类
' 输入：    无
' 返回：    无
'*********************************************************
Class TAtom10Export

	Public TimeZone

	Public Property Get xml
		xml = objXMLdoc.xml
	End Property

	Private objXMLdoc
	Private objXMLfeed

	Public Function GetFeed(objAtomFeed)

		Set objXMLfeed=objAtomFeed
		objXMLdoc.AppendChild(objXMLfeed)
		objXMLfeed.setAttribute "xmlns","http://www.w3.org/2005/Atom"

		Dim i
		Dim objItemNodes
		Set objItemNodes=objXMLfeed.getElementsByTagName("updated")

		For i=0 To (objItemNodes.Length-1)
			objItemNodes(i).Text=ParseDateForRFC3339(objItemNodes(i).Text)
		Next

		Set objItemNodes=Nothing

	End Function


	Public Function GetEntry(objEntryFeed)

		Dim i
		Dim objItemNodes
		Set objItemNodes=objEntryFeed.getElementsByTagName("updated")

		For i=0 To (objItemNodes.Length-1)
			objItemNodes(i).Text=ParseDateForRFC3339(objItemNodes(i).Text)
		Next

		Set objItemNodes=Nothing

		Set objItemNodes=objEntryFeed.getElementsByTagName("published")

		For i=0 To (objItemNodes.Length-1)
			objItemNodes(i).Text=ParseDateForRFC3339(objItemNodes(i).Text)
		Next

		Set objItemNodes=Nothing

		objXMLfeed.appendChild(objEntryFeed)

	End Function


	Public Function Execute()

		'Response.ContentType = "text/html"
		Response.ContentType = "text/xml"
		Response.Clear
		Response.Write xml

		Execute=True

	End Function


	Public Function SaveToFile(strFileName)

		objXMLdoc.save(strFileName)

		SaveToFile=True

	End Function


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

		strTimeZone=Left(TimeZone,3) & ":" & Right(TimeZone,2)

		ParseDateForRFC3339 = dtmYear & "-" & dtmMonth & "-" & dtmDay & "T" & dtmHours & ":" & dtmMinutes & ":" & dtmSeconds & strTimeZone

	End Function 

	' 类初始化
	Private Sub Class_Initialize()

		On Error Resume Next

		'对objXMLdoc进行初始化，如不能建对象则报错
		Set objXMLdoc =Server.CreateObject("Microsoft.XMLDOM")

		If Err.Number<>0 Then

		End If

		Dim objPI

		Set objPI = objXMLdoc.createProcessingInstruction("xml-stylesheet","type=""text/css"" href=""css/atom.css""")
		objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
		Set objPI = Nothing

		Set objPI = objXMLdoc.createProcessingInstruction("xml","version=""1.0"" encoding=""UTF-8"" standalone=""yes""")
		objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
		Set objPI = Nothing

		Set objXMLfeed = objXMLdoc.createElement("feed")

	End Sub

	' 类释放
	Private Sub Class_Terminate()

		Set objXMLfeed = Nothing
		Set objXMLdoc  = Nothing

	End Sub

End Class
'*********************************************************




'*********************************************************
' 目的：    BLOG信息类
' 输入：    无
' 返回：    无
'*********************************************************
Class TAtomFeed

	Public Property Get Node
		Set Node=objFeedNode
	End Property

	'Public atomAuthor
	'Public atomCategory
	'Public atomContributor
	'Public atomGenerator
	'Public atomIcon
	'Public atomId
	'Public atomLink
	'Public atomLogo
	'Public atomRights
	'Public atomSubtitle
	'Public atomTitle
	'Public atomUpdated

	Private objXMLdoc
	Private objFeedNode


	Private Function CommomAppendNode(strElement,strText,strType)

		Dim objSingleNode
		Dim objNodeText
		Dim objNodeCdata
		Set objSingleNode = objXMLdoc.createNode("element",strElement,"")

		If strType="" Then

			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		ElseIf strType="text" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		ElseIf strType="html" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeCdata=objXMLdoc.createNode("cdatasection", "", "")
			objNodeCdata.NodeValue=strText
			objSingleNode.AppendChild(objNodeCdata)

		ElseIf strType="xhtml" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		End If

		Set CommomAppendNode=objSingleNode

		Set objSingleNode = Nothing

	End Function

	Public Property Let atomCategory(strCategory)

		objFeedNode.AppendChild(CommomAppendNode("category",strCategory,""))

	End Property


	Public Property Let atomIcon(strIcon)

		objFeedNode.AppendChild(CommomAppendNode("icon",strIcon,""))

	End Property


	Public Property Let atomId(strId)

		objFeedNode.AppendChild(CommomAppendNode("id",strId,""))

	End Property


	Public Property Let atomLogo(strLogo)

		objFeedNode.AppendChild(CommomAppendNode("logo",strLogo,""))

	End Property


	Public Property Let atomRights(strRights)

		objFeedNode.AppendChild(CommomAppendNode("rights",strRights,"text"))

	End Property


	Public Property Let atomSubtitle(strSubtitle)

		objFeedNode.AppendChild(CommomAppendNode("subtitle",strSubtitle,"html"))

	End Property


	Public Property Let atomTitle(strTitle)

		objFeedNode.AppendChild(CommomAppendNode("title",strTitle,"html"))

	End Property


	Public Property Let atomUpdated(strUpdated)

		objFeedNode.AppendChild(CommomAppendNode("updated",strUpdated,""))

	End Property


	Public Function atomPerson(strPerson,strName,strEmail,strUrl)

		Dim objSingleNode
		Dim objNodeText

		Dim objAuthorNameNode
		Dim objAuthorUrlNode
		Dim objAuthorEmailNode

		Set objSingleNode = objXMLdoc.createNode("element",strPerson,"")

		If strName<>"" Then
			Set objAuthorNameNode = objXMLdoc.createNode("element","name","")
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strName
			objAuthorNameNode.AppendChild(objNodeText)
			objSingleNode.AppendChild(objAuthorNameNode)
			Set objNodeText = Nothing
		End If

		If strUrl<>"" Then
			Set objAuthorUrlNode = objXMLdoc.createNode("element","uri","")
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strUrl
			objAuthorUrlNode.AppendChild(objNodeText)
			objSingleNode.AppendChild(objAuthorUrlNode)
			Set objNodeText = Nothing
		End If

		If strEmail<>"" Then
			Set objAuthorEmailNode = objXMLdoc.createNode("element","email","")
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strEmail
			objAuthorEmailNode.AppendChild(objNodeText)
			objSingleNode.AppendChild(objAuthorEmailNode)
			Set objNodeText = Nothing
		End If

		objFeedNode.AppendChild(objSingleNode)

		Set objSingleNode = Nothing

	End Function


	Public Function atomLink(strRel,strType,strHref)

		Dim objSingleNode
		Dim objNodeText

		Set objSingleNode = objXMLdoc.createNode("element","link","")

		objSingleNode.setAttribute "rel",strRel
		objSingleNode.setAttribute "type",strType
		objSingleNode.setAttribute "href",strHref

		objFeedNode.AppendChild(objSingleNode)
		Set objSingleNode = Nothing

	End Function


	Public Function atomGenerator(strGenerator,strUri,strVersion)

		Dim objSingleNode
		Dim objNodeText

		Set objSingleNode = objXMLdoc.createNode("element","generator","")
		Set objNodeText=objXMLdoc.createNode("text", "", "")

		objNodeText.NodeValue=strGenerator
		objSingleNode.setAttribute "uri",strUri
		objSingleNode.setAttribute "version",strVersion

		objSingleNode.AppendChild(objNodeText)
		objFeedNode.AppendChild(objSingleNode)

		Set objSingleNode = Nothing
		Set objNodeText = Nothing

	End Function


	Private Sub Class_Initialize()

		Set objXMLdoc =Server.CreateObject("Microsoft.XMLDOM")
		Set objFeedNode = objXMLdoc.createElement("feed")

	End Sub


	Private Sub Class_Terminate()

		Set objXMLdoc = Nothing
		Set objFeedNode = Nothing

	End Sub


End Class
'*********************************************************




'*********************************************************
' 目的：    日志类
' 输入：    无
' 返回：    无
'*********************************************************
Class TAtomEntry

	Public Property Get Node
		Set Node=objEntryNode
	End Property

	'Public atomAuthor
	'Public atomCategory
	'Public atomContributor
	'Public atomLink
	'Public atomTitle
	'Public atomUpdated
	'Public atomPublished
	'Public atomContent
	'Public atomSummary
	'Public atomId
	'Public atomRights

	Private objXMLdoc
	Private objEntryNode

	Private Function CommomAppendNode(strElement,strText,strType)

		Dim objSingleNode
		Dim objNodeText
		Dim objNodeCdata
		Set objSingleNode = objXMLdoc.createNode("element",strElement,"")

		If strType="" Then

			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		ElseIf strType="text" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		ElseIf strType="html" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeCdata=objXMLdoc.createNode("cdatasection", "", "")
			objNodeCdata.NodeValue=strText
			objSingleNode.AppendChild(objNodeCdata)

		ElseIf strType="xhtml" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		End If

		Set CommomAppendNode=objSingleNode

		Set objSingleNode = Nothing

	End Function


	Public Function atomContent(strContent,strType)

		objEntryNode.AppendChild(CommomAppendNode("content",strContent,strType))

	End Function


	Public Property Let atomSummary(strSummary)

		objEntryNode.AppendChild(CommomAppendNode("summary",strSummary,"html"))

	End Property


	Public Property Let atomRights(strRights)

		objEntryNode.AppendChild(CommomAppendNode("rights",strRights,""))

	End Property


	Public Property Let atomId(strID)

		objEntryNode.AppendChild(CommomAppendNode("id",strID,""))

	End Property


	Public Property Let atomUpdated(dtmUpdated)

		objEntryNode.AppendChild(CommomAppendNode("updated",dtmUpdated,""))

	End Property


	Public Property Let atomPublished(dtmPublished)

		objEntryNode.AppendChild(CommomAppendNode("published",dtmPublished,""))

	End Property


	Public Property Let atomTitle(strTitle)

		objEntryNode.AppendChild(CommomAppendNode("title",strTitle,"html"))

	End Property


	Public Function atomCategory(strTerm,strScheme,strLabel)

		Dim objSingleNode

		Set objSingleNode = objXMLdoc.createNode("element","category","")

		objSingleNode.setAttribute "term",strTerm
		objSingleNode.setAttribute "scheme",strScheme
		objSingleNode.setAttribute "label",strLabel

		objEntryNode.AppendChild(objSingleNode)
		Set objSingleNode = Nothing

	End Function


	Public Function atomPerson(strPerson,strName,strEmail,strUrl)

		Dim objSingleNode
		Dim objNodeText

		Dim objAuthorNameNode
		Dim objAuthorUrlNode
		Dim objAuthorEmailNode

		Set objSingleNode = objXMLdoc.createNode("element",strPerson,"")

		If strName<>"" Then
			Set objAuthorNameNode = objXMLdoc.createNode("element","name","")
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strName
			objAuthorNameNode.AppendChild(objNodeText)
			objSingleNode.AppendChild(objAuthorNameNode)
			Set objNodeText = Nothing
		End If

		If strUrl<>"" Then
			Set objAuthorUrlNode = objXMLdoc.createNode("element","uri","")
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strUrl
			objAuthorUrlNode.AppendChild(objNodeText)
			objSingleNode.AppendChild(objAuthorUrlNode)
			Set objNodeText = Nothing
		End If

		If strEmail<>"" Then
			Set objAuthorEmailNode = objXMLdoc.createNode("element","email","")
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strEmail
			objAuthorEmailNode.AppendChild(objNodeText)
			objSingleNode.AppendChild(objAuthorEmailNode)
			Set objNodeText = Nothing
		End If

		objEntryNode.AppendChild(objSingleNode)

		Set objSingleNode = Nothing

	End Function

	Public Function atomLink(strRel,strType,strHref)

		Dim objSingleNode
		Dim objNodeText

		Set objSingleNode = objXMLdoc.createNode("element","link","")

		objSingleNode.setAttribute "rel",strRel
		objSingleNode.setAttribute "type",strType
		objSingleNode.setAttribute "href",strHref

		objEntryNode.AppendChild(objSingleNode)
		Set objSingleNode = Nothing

	End Function


	Public Property Let atomTag(strTag)

		objEntryNode.AppendChild(CommomAppendNode("tag",strTag,""))

	End Property


	Public Function GetComment(objCommentFeed)

		objEntryNode.appendChild(objCommentFeed)

	End Function


	Private Sub Class_Initialize()

		Set objXMLdoc =Server.CreateObject("Microsoft.XMLDOM")
		Set objEntryNode = objXMLdoc.createElement("entry")

	End Sub


	Private Sub Class_Terminate()

		Set objXMLdoc = Nothing
		Set objEntryNode = Nothing

	End Sub


End Class
'*********************************************************




'*********************************************************
' 目的：    评论类
' 输入：    无
' 返回：    无
'*********************************************************
Class TAtomComment

	Public Property Get Node
		Set Node=objCommentNode
	End Property

	'Public atomAuthor
	'Public atomPublished
	'Public atomContent

	Private objXMLdoc
	Private objCommentNode

	Private Function CommomAppendNode(strElement,strText,strType)

		Dim objSingleNode
		Dim objNodeText
		Dim objNodeCdata
		Set objSingleNode = objXMLdoc.createNode("element",strElement,"")

		If strType="" Then

			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		ElseIf strType="text" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		ElseIf strType="html" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeCdata=objXMLdoc.createNode("cdatasection", "", "")
			objNodeCdata.NodeValue=strText
			objSingleNode.AppendChild(objNodeCdata)

		ElseIf strType="xhtml" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		End If

		Set CommomAppendNode=objSingleNode

		Set objSingleNode = Nothing

	End Function


	Public Property Let atomTitle(strTitle)

		objCommentNode.AppendChild(CommomAppendNode("title",strTitle,"html"))

	End Property


	Public Function atomContent(strContent,strType)

		objCommentNode.AppendChild(CommomAppendNode("content",strContent,strType))

	End Function


	Public Property Let atomPublished(dtmPublished)

		objCommentNode.AppendChild(CommomAppendNode("published",dtmPublished,""))

	End Property


	Public Function atomPerson(strPerson,strName,strEmail,strUrl)

		Dim objSingleNode
		Dim objNodeText

		Dim objAuthorNameNode
		Dim objAuthorUrlNode
		Dim objAuthorEmailNode

		Set objSingleNode = objXMLdoc.createNode("element",strPerson,"")

		If strName<>"" Then
			Set objAuthorNameNode = objXMLdoc.createNode("element","name","")
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strName
			objAuthorNameNode.AppendChild(objNodeText)
			objSingleNode.AppendChild(objAuthorNameNode)
			Set objNodeText = Nothing
		End If

		If strUrl<>"" Then
			Set objAuthorUrlNode = objXMLdoc.createNode("element","uri","")
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strUrl
			objAuthorUrlNode.AppendChild(objNodeText)
			objSingleNode.AppendChild(objAuthorUrlNode)
			Set objNodeText = Nothing
		End If

		If strEmail<>"" Then
			Set objAuthorEmailNode = objXMLdoc.createNode("element","email","")
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strEmail
			objAuthorEmailNode.AppendChild(objNodeText)
			objSingleNode.AppendChild(objAuthorEmailNode)
			Set objNodeText = Nothing
		End If

		objCommentNode.AppendChild(objSingleNode)

		Set objSingleNode = Nothing

	End Function


	Private Sub Class_Initialize()

		Set objXMLdoc =Server.CreateObject("Microsoft.XMLDOM")
		Set objCommentNode = objXMLdoc.createElement("comment")

	End Sub


	Private Sub Class_Terminate()

		Set objXMLdoc = Nothing
		Set objCommentNode = Nothing

	End Sub


End Class
'*********************************************************




'*********************************************************
' 目的：    评论类
' 输入：    无
' 返回：    无
'*********************************************************
Class TAtomTrackBack

	Public Property Get Node
		Set Node=objTrackBackNode
	End Property

	'Public atomAuthor
	'Public atomPublished
	'Public atomContent

	Private objXMLdoc
	Private objTrackBackNode

	Private Function CommomAppendNode(strElement,strText,strType)

		Dim objSingleNode
		Dim objNodeText
		Dim objNodeCdata
		Set objSingleNode = objXMLdoc.createNode("element",strElement,"")

		If strType="" Then

			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		ElseIf strType="text" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		ElseIf strType="html" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeCdata=objXMLdoc.createNode("cdatasection", "", "")
			objNodeCdata.NodeValue=strText
			objSingleNode.AppendChild(objNodeCdata)

		ElseIf strType="xhtml" Then

			objSingleNode.setAttribute "type",strType
			Set objNodeText=objXMLdoc.createNode("text", "", "")
			objNodeText.NodeValue=strText
			objSingleNode.AppendChild(objNodeText)

		End If

		Set CommomAppendNode=objSingleNode

		Set objSingleNode = Nothing

	End Function


	Public Property Let atomTitle(strTitle)

		objTrackBackNode.AppendChild(CommomAppendNode("title",strTitle,"html"))

	End Property


	Public Function atomContent(strContent,strType)

		objTrackBackNode.AppendChild(CommomAppendNode("content",strContent,strType))

	End Function


	Public Property Let atomPublished(dtmPublished)

		objTrackBackNode.AppendChild(CommomAppendNode("published",dtmPublished,""))

	End Property


	Public Function atomLink(strRel,strType,strHref)

		Dim objSingleNode
		Dim objNodeText

		Set objSingleNode = objXMLdoc.createNode("element","link","")

		objSingleNode.setAttribute "rel",strRel
		objSingleNode.setAttribute "type",strType
		objSingleNode.setAttribute "href",strHref

		objTrackBackNode.AppendChild(objSingleNode)
		Set objSingleNode = Nothing

	End Function


	Private Sub Class_Initialize()

		Set objXMLdoc =Server.CreateObject("Microsoft.XMLDOM")
		Set objTrackBackNode = objXMLdoc.createElement("trackback")

	End Sub


	Private Sub Class_Terminate()

		Set objXMLdoc = Nothing
		Set objTrackBackNode = Nothing

	End Sub


End Class
'*********************************************************

</Script>