<%

	Dim xmlPost, ToUserName, FromUserName, CreateTime, MsgType, Content, MsgId
	'xmlPost = "<xml><ToUserName><![CDATA[imzhoujie]]></ToUserName> <FromUserName><![CDATA[fasongren]]></FromUserName><CreateTime>1348831860</CreateTime> <MsgType><![CDATA[text]]></MsgType> <Content><![CDATA[this is a test]]></Content> <MsgId>1234567890123456</MsgId> </xml>"
	xmlPost = Request.InputStream
	Dim xmlReader, record	'从xml字符串中读取
	Set xmlReader = Server.CreateObject("msxml.domdocument")
	xmlReader.loadXML(xmlPost)
	'Set record = xmlReader.documentElement.selectNodes("//")
	'Response.Write("Record个数:" & record.length & "<br/>")
	
	'内容
	ToUserName = xmlReader.documentElement.selectNodes("//ToUserName")(0).firstChild.nodeValue
	FromUserName = xmlReader.documentElement.selectNodes("//FromUserName")(0).firstChild.nodeValue
	CreateTime = xmlReader.documentElement.selectNodes("//CreateTime")(0).firstChild.nodeValue
	MsgType = xmlReader.documentElement.selectNodes("//MsgType")(0).firstChild.nodeValue
	Content = xmlReader.documentElement.selectNodes("//Content")(0).firstChild.nodeValue
	MsgId = xmlReader.documentElement.selectNodes("//MsgId")(0).firstChild.nodeValue
	Set xmlReader = Nothing

	reply="<xml><ToUserName><![CDATA[{ToUserName}]]></ToUserName><FromUserName><![CDATA[{FromUserName}]]></FromUserName><CreateTime>{CreateTime}</CreateTime><MsgType><![CDATA[{MsgType}]]></MsgType><Content><![CDATA[{Content}]]></Content><FuncFlag>{FuncFlag}</FuncFlag></xml>"

	reply=Replace(reply,"{ToUserName}",FromUserName)
	reply=Replace(reply,"{FromUserName}","imzhoujie")
	reply=Replace(reply,"{CreateTime}",CreateTime+2)
	reply=Replace(reply,"{MsgType}","text")
	reply=Replace(reply,"{Content}","未寒博客自动回复")
	reply=Replace(reply,"{FuncFlag}","0")

	response.write reply
	Response.End()
%>