<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../zb_system/function/c_function.asp" -->
<!-- #include file="../../../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../../../zb_system/function/c_system_base.asp" -->
<!-- #include file="../../../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../p_config.asp" -->
<%
Call System_Initialize()
If CheckPluginState("weixin")=False Then Response.Write "插件未启用":Response.End()
'Response.ContentType="application/xml"

Dim objConfig
Set objConfig=New TConfig
objConfig.Load("weixin")

If Request("echostr") Then
	'Dim echostr,signature,timestamp,nonce,token,tmpArray,tmpStr,i,singsha1
	'signature = Request("signature")
	'timestamp = Request("timestamp")
	'nonce = Request("nonce")
	echostr = Request("echostr")
	'tmpArr = Array(nonce,timestamp,objConfig.Read("token"))
	'tmpArr = Sort(tmpArr)
	'tmpStr = ""
	'For i = 0 to 2
    '  tmpStr=tmpStr & tmpArr(i)
	'Next
	'singsha1 = hex_sha1(tmpStr)
	'If(singsha1 = signature) Then
		Response.Write echostr
		Response.End()
	'else
	'	Response.End()
	'End If
End If

Dim ToUserName	'开发者微信号
Dim FromUserName'发送方帐号（一个OpenID）
Dim CreateTime	'消息创建时间（整型）
Dim MsgType		'text
Dim Content		'文本消息内容
Dim MsgId		'消息id，64位整型
Dim varEvent,varEventKey	'事件消息
Dim strresponse_text,strresponse_news

Dim xml_dom
set xml_dom = Server.CreateObject("MSXML2.DOMDocument")
xml_dom.load Request
ToUserName=xml_dom.getelementsbytagname("ToUserName").item(0).text
FromUserName=xml_dom.getelementsbytagname("FromUserName").item(0).text
MsgType=xml_dom.getelementsbytagname("MsgType").item(0).text
if MsgType="event" then
	varEvent=xml_dom.getelementsbytagname("Event").item(0).text
	if varEvent="subscribe" then
		Content=wx_Welcome(ZC_BLOG_TITLE,objConfig.Read("LastPostNum"),objConfig.Read("WelcomeStr"))
		strresponse_text="<xml>" &_
		"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
		"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
		"<CreateTime>"&now&"</CreateTime>" &_
		"<MsgType><![CDATA[text]]></MsgType>" &_
		"<Content><![CDATA[" & Content & "]]></Content>" &_
		"<FuncFlag>0<FuncFlag>" &_
		"</xml>"
		response.write strresponse_text
		Response.End()
	end if
elseif MsgType="text" then
	Content=xml_dom.getelementsbytagname("Content").item(0).text
end if
set xml_dom=Nothing

Dim strQuestion
strQuestion=Trim(TransferHTML(Content,"[nohtml]"))
'strQuestion=FilterSQL(strQuestion)
strQuestion=Replace(strQuestion," ","")
	strQuestion=Replace(strQuestion,"'","")
	strQuestion=Replace(strQuestion,"“","")
	strQuestion=Replace(strQuestion,"”","")
	strQuestion=Replace(strQuestion,"‘","")
	strQuestion=Replace(strQuestion,"’","")
	strQuestion=Replace(strQuestion,"AND","")
	strQuestion=Replace(strQuestion,"WHERE","")
If (strQuestion="help") Then
	Content = wx_Help()
	strresponse_text="<xml>" &_
 	"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
	"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
	"<CreateTime>"&now&"</CreateTime>" &_
	"<MsgType><![CDATA[text]]></MsgType>" &_
	"<Content><![CDATA[" & Content & "]]></Content>" &_
	"<FuncFlag>0<FuncFlag>" &_
	"</xml>"
	response.write strresponse_text
ElseIf (strQuestion="最新文章") Then
	Content = wx_LastPost(CInt(objConfig.Read("LastPostNum")))
	strresponse_news="<xml>"&_
	"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>"&_
	"<FromUserName><![CDATA["&tousername&"]]></FromUserName>"&_
	"<CreateTime>"&now&"</CreateTime>"&_
	"<MsgType><![CDATA[news]]></MsgType>"&_
	"<ArticleCount>"&CInt(objConfig.Read("LastPostNum"))&"</ArticleCount>"&_
	"<Articles>"& Content &_	
	"</Articles>"&_
	"<FuncFlag>1</FuncFlag>"&_
	"</xml>"
	response.write strresponse_news
	Response.End()
Else
	Content = wx_Search(strQuestion,CInt(objConfig.Read("SearchNum")),CInt(objConfig.Read("ShowMeta")))
	strresponse_text="<xml>" &_
 	"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
	"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
	"<CreateTime>"&now&"</CreateTime>" &_
	"<MsgType><![CDATA[text]]></MsgType>" &_
	"<Content><![CDATA[" & Content & "]]></Content>" &_
	"<FuncFlag>0<FuncFlag>" &_
	"</xml>"
	response.write strresponse_text
End If

'If FromUserName = "" Then Response.End()
%>