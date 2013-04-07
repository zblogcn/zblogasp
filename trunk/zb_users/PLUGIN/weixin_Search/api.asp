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
Response.ContentType="application/x-javascript"
Call System_Initialize()

If Request("echostr") Then
	Dim echostr,signature,timestamp,nonce,token,tmpArray,tmpStr,i
	signature = Request("signature")
	timestamp = Request("timestamp")
	nonce = Request("nonce")
	echostr = Request("echostr")
	'Array.Sort(ArrTmp)
	tmpArr = Array(nonce,timestamp,"imzhou")
	tmpArr = Sort(tmpArr)
	tmpStr = ""
	For i = 0 to 2
      tmpStr=tmpStr & tmpArr(i)
	Next
	singsha1 = hex_sha1(tmpStr)
	if(singsha1 = signature) then
		response.write echostr
		Response.End()
	end if
End If


Dim ToUserName	'开发者微信号
Dim FromUserName'发送方帐号（一个OpenID）
Dim CreateTime	'消息创建时间（整型）
Dim MsgType		'text
Dim Content		'文本消息内容
Dim MsgId		'消息id，64位整型
Dim xml_dom,strresponse

set xml_dom = Server.CreateObject("MSXML2.DOMDocument")'此处根据您的实际服务器情况改写
xml_dom.load Request
ToUserName=xml_dom.getelementsbytagname("ToUserName").item(0).text
FromUserName=xml_dom.getelementsbytagname("FromUserName").item(0).text
MsgType=xml_dom.getelementsbytagname("MsgType").item(0).text
if MsgType="text" then
Content=xml_dom.getelementsbytagname("Content").item(0).text
end if
set xml_dom=Nothing

	'dim filepath,fso,fopen
	'filepath=server.mappath(".")&"\wx.txt"
	'Set fso = Server.CreateObject("Scripting.FileSystemObject")
	'set fopen=fso.OpenTextFile(filepath, 8 ,true)
	'fopen.writeline(Request)
	'set fso=nothing
	'set fopen=Nothing
	
'Content="花"
Dim strQuestion
strQuestion=TransferHTML(Content,"[nohtml]")

	dim LTRS,InserNewHtml:InserNewHtml = ""
	Set LTRS=objConn.Execute("SELECT [log_ID],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_PostTime],[log_Url],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE ([log_Type]=0) And ([log_ID]>0) AND( (InStr(1,LCase([log_Title]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Intro]),LCase('"&strQuestion&"'),0)<>0) OR (InStr(1,LCase([log_Content]),LCase('"&strQuestion&"'),0)<>0) )")
	Do Until LTRS.Eof
		InserNewHtml = InserNewHtml & LTRS("log_Title") & "----" & LTRS("log_PostTime") & VBCrLf
		InserNewHtml = InserNewHtml & TransferHTML(LTRS("log_Content"),"[nohtml]")
		Exit Do
		'LTRS.MoveNext
	Loop
	Set LTRS=Nothing

	Content = InserNewHtml
	Content = Replace(Content,"&nbsp;"," ")
	'Content = Replace(Content,"<#ZC_BLOG_HOST#>",BlogHost)
	
strresponse="<xml>" &_
 	"<ToUserName><![CDATA["&fromusername&"]]></ToUserName>" &_
	"<FromUserName><![CDATA["&tousername&"]]></FromUserName>" &_
	"<CreateTime>"&now&"</CreateTime>" &_
	"<MsgType><![CDATA[text]]></MsgType>" &_
	"<Content><![CDATA[" & Content & "]]></Content>" &_
	"<FuncFlag>0<FuncFlag>" &_
	"</xml>"
response.write strresponse


%>
