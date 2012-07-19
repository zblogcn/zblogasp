<%@ CODEPAGE=65001 %>
<FONT SIZE="" COLOR=""></FONT><!--#include file="up_inc.asp"-->
<!-- #include file="../../../../zb_users\c_option.asp" -->
<!-- #include file="../../../function\c_function.asp" -->
<!-- #include file="../../../function\c_system_lib.asp" -->
<!-- #include file="../../../function\c_system_base.asp" -->
<!-- #include file="../../../function\c_system_event.asp" -->
<!-- #include file="../../../function\c_system_plugin.asp" -->
<!-- #include file="../../../../zb_users\plugin\p_config.asp" -->
<%
On Error Resume Next
Call System_Initialize()
Call CheckReference("")
If Not CheckRights("ArticleEdt") Then Call ShowError(6)

For Each sAction_Plugin_getmovie_Begin in Action_Plugin_getmovie_Begin
	If Not IsEmpty(sAction_Plugin_getmovie_Begin) Then Call Execute(sAction_Plugin_getmovie_Begin)
Next
	Dim strResponse
	'strResponse="此功能(getmovie.asp)系统默认不开放，请安装必要插件。"
	Dim key,type2
	key=Trim(Request.Form("searchKey"))
	type2=Trim(Request.Form("videoType"))
	strResponse=gethtml("http://api.tudou.com/v3/gw?method=item.search&appKey=myKey&format=json&kw="&key&"&pageNo=1&pageSize=20&channelId="&type2&"&inDays=7&media=v&sort=s")

For Each sAction_Plugin_getmovie_End in Action_Plugin_getmovie_End
	If Not IsEmpty(sAction_Plugin_getmovie_End) Then Call Execute(sAction_Plugin_getmovie_End)
Next
	Response.Write strResponse
Call System_Terminate()
%>

<%




function gethtml(strUrl)
	on error resume next
	dim objXmlHttp
	set objXmlHttp=server.createobject("MSXML2.ServerXMLHTTP")
	objXmlHttp.setTimeouts 10000,10000,10000,30000
	objXmlHttp.open "GET",strUrl,false
	objXmlHttp.send()
	gethtml=BytesToBstr( objXmlHttp.responseBody,"utf-8")
	err.Clear
	set objXmlHttp=nothing
end function

Function BytesToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText
objstream.Close
set objstream = nothing
End Function

%>