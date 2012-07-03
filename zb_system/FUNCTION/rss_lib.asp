<%
'///////////////////////////////////////////////////////////////////////////////
'//              RainbowSoft RSS Export
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    rss_lib.asp
'// 开始时间:    2004.07.03
'// 最后修改:    
'// 备    注:    RSS输出
'///////////////////////////////////////////////////////////////////////////////




'*********************************************************
' 目的：    定义TRss2Export类
' 输入：    无
' 返回：    无
'*********************************************************
Class TRss2Export

	Public TimeZone

	Public Property Get xml
		xml = objXMLdoc.xml
	End Property

	public FstrWebLink
	public FstrAuthor

	Public Property Get WebLink
		WebLink = FstrWebLink
	End Property

	Public Property Let WebLink(strWebLink)
		FstrWebLink = strWebLink
	End Property

	Public Property Get Author
		Author = FstrAuthor
	End Property

	Public Property Let Author(strAuthor)
		FstrAuthor = strAuthor
	End Property

	Private objXMLdoc

	Private objXMLrss

	Private objXMLchannel


	Public Function AddChannelAttribute(title,value)

		Dim objXMLitem
		Set objXMLitem = objXMLdoc.createElement(title)

		If title="pubDate" Then value=ParseDateForRFC822(value)

		objXMLitem.text=value
		objXMLchannel.AppendChild(objXMLitem)

		AddChannelAttribute=True

	End Function


	Public Function AddItem(title,author,link,pubDate,guid,description,category,comments,wfw_comment,wfw_commentRss,trackback_ping)

		Dim objXMLitem
		Set objXMLitem = objXMLdoc.createElement("item")
		Dim objXMLcdata

		If(Len(title)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("title"))
			objXMLitem.selectSingleNode("title").text=title
		End If
		If(Len(author)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("author"))
			objXMLitem.selectSingleNode("author").text=author
		End If
		If(Len(link)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("link"))
			objXMLitem.selectSingleNode("link").text=link
		End If
		If(Len(pubDate)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("pubDate"))
			objXMLitem.selectSingleNode("pubDate").text=ParseDateForRFC822(pubDate)
		End If
		If(Len(guid)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("guid"))
			objXMLitem.selectSingleNode("guid").text=guid
		End If
		If(Len(description)>0) Then

			objXMLitem.AppendChild(objXMLdoc.createElement("description"))
			Set objXMLcdata = objXMLdoc.createNode("cdatasection", "","")
			objXMLcdata.NodeValue=description
			objXMLitem.selectSingleNode("description").AppendChild(objXMLcdata)

			Set objXMLcdata = Nothing

		End If
		If(Len(category)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("category"))
			objXMLitem.selectSingleNode("category").text=category
		End If

		If(Len(comments)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("comments"))
			objXMLitem.selectSingleNode("comments").text=comments
		End If
		If(Len(wfw_comment)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("wfw:comment"))
			objXMLitem.selectSingleNode("wfw:comment").text=wfw_comment
		End If
		If(Len(wfw_commentRss)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("wfw:commentRss"))
			objXMLitem.selectSingleNode("wfw:commentRss").text=wfw_commentRss
		End If
		If(Len(trackback_ping)>0) Then
			objXMLitem.AppendChild(objXMLdoc.createElement("trackback:ping"))
			objXMLitem.selectSingleNode("trackback:ping").text=trackback_ping
		End If

		objXMLchannel.AppendChild(objXMLitem)

		AddItem=True

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


	Function ParseDateForRFC822(dtmDate)

		Dim dtmDay, dtmWeekDay, dtmMonth, dtmYear
		Dim dtmHours, dtmMinutes, dtmSeconds

		Select Case WeekDay(dtmDate)
			Case 1:dtmWeekDay="Sun"
			Case 2:dtmWeekDay="Mon"
			Case 3:dtmWeekDay="Tue"
			Case 4:dtmWeekDay="Wed"
			Case 5:dtmWeekDay="Thu"
			Case 6:dtmWeekDay="Fri"
			Case 7:dtmWeekDay="Sat"
		End Select

		Select Case Month(dtmDate)
			Case 1:dtmMonth="Jan"
			Case 2:dtmMonth="Feb"
			Case 3:dtmMonth="Mar"
			Case 4:dtmMonth="Apr"
			Case 5:dtmMonth="May"
			Case 6:dtmMonth="Jun"
			Case 7:dtmMonth="Jul"
			Case 8:dtmMonth="Aug"
			Case 9:dtmMonth="Sep"
			Case 10:dtmMonth="Oct"
			Case 11:dtmMonth="Nov"
			Case 12:dtmMonth="Dec"
		End Select

		dtmYear = Year(dtmDate)
		dtmDay = Right("00" & Day(dtmDate),2)

		dtmHours = Right("00" & Hour(dtmDate),2)
		dtmMinutes = Right("00" & Minute(dtmDate),2)
		dtmSeconds = Right("00" & Second(dtmDate),2)

		ParseDateForRFC822 = dtmWeekDay & ", " & dtmDay &" " & dtmMonth & " " & dtmYear & " " & dtmHours & ":" & dtmMinutes & ":" & dtmSeconds & " " & TimeZone

	End Function 

	' 类初始化
	Private Sub Class_Initialize()

		On Error Resume Next

		'对objXMLdoc进行初始化，如不能建对象则报错
		Set objXMLdoc =Server.CreateObject("Microsoft.XMLDOM")

		If Err.Number<>0 Then

		End If

		Dim objPI

		'Set objPI = objXMLdoc.createProcessingInstruction("xml-stylesheet","type=""text/css"" href=""css/rss.css""")
		'objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
		'Set objPI = Nothing

		Set objPI = objXMLdoc.createProcessingInstruction("xml-stylesheet","type=""text/xsl"" href=""css/rss.xslt""")
		objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
		Set objPI = Nothing

		Set objPI = objXMLdoc.createProcessingInstruction("xml","version=""1.0"" encoding=""UTF-8"" standalone=""yes""")
		objXMLdoc.insertBefore objPI, objXMLdoc.childNodes(0)
		Set objPI = Nothing

		Set objXMLrss = objXMLdoc.createElement("rss")

		Set objXMLchannel = objXMLdoc.createElement("channel")


		objXMLrss.AppendChild(objXMLchannel)
		objXMLdoc.AppendChild(objXMLrss)

		objXMLrss.setAttribute "version","2.0"
		objXMLrss.setAttribute "xmlns:dc","http://purl.org/dc/elements/1.1/"
		objXMLrss.setAttribute "xmlns:trackback","http://madskills.com/public/xml/rss/module/trackback/"
		objXMLrss.setAttribute "xmlns:wfw","http://wellformedweb.org/CommentAPI/"
		objXMLrss.setAttribute "xmlns:slash","http://purl.org/rss/1.0/modules/slash/"


	End Sub

	' 类释放
	Private Sub Class_Terminate()

		Set objXMLrss = Nothing
		Set objXMLdoc  = Nothing

	End Sub

End Class
'*********************************************************

%>