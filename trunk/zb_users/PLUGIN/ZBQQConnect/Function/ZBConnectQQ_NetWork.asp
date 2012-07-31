<%
Class ZBQQConnect_NetWork
Dim objXmlhttp
'Dim objWinhttp
Public Status,ReadyState
Public ResponseText,ResponseBody
Public CharSet
Public Par
Public UA

Sub Class_Initialize
	Set objXmlhttp=Server.CreateObject("msxml2.serverxmlhttp")
	CharSet="utf-8"
	UA="ZSXSOFT"
	Set Par=ZBQQConnect_Toobject("{}")
End Sub

Public Sub setRequestHeader()
	Dim a
	a=ZBQQConnect_ToStr(Par)
	Dim b,c,d
	b=Split(a,"&")
	For c=0 to Ubound(b)
		d=Split(b(c),"=")
		objXmlhttp.setRequestHeader d(0),d(1)
	Next
End Sub

Public Function GetHttp(Url)
	objXmlhttp.SetTimeOuts 10000, 10000, 10000, 10000 
	objXmlhttp.Open "GET",url
	Call ZBQQConnect_addObj(Par,"User-Agent",UA)
	setRequestHeader
	objXmlhttp.Send
	ResponseText=objXmlhttp.ResponseText
	ResponseBody=objXmlhttp.ResponseBody
	GetHttp=BytesToBstr(ResponseBody,CharSet)
	Set Par=ZBQQConnect_Toobject("{}")
End Function

Public Function PostHttp(Url,Data)
	objXmlhttp.SetTimeOuts 10000, 10000, 10000, 10000 
	objXmlhttp.Open "POST",url
	Call ZBQQConnect_addObj(Par,"Content-type","application/x-www-form-urlencoded")
	Call ZBQQConnect_addObj(Par,"User-Agent",UA)
	setRequestHeader
	objXmlhttp.Send Data
	ResponseText=objXmlhttp.ResponseText
	ResponseBody=objXmlhttp.ResponseBody
	PostHttp=BytesToBstr(ResponseBody,CharSet)
	Set Par=ZBQQConnect_Toobject("{}")
End Function


Function BytesToBstr(body,Cset)
	dim objstream
	set objstream=createobject("adodb.stream")
	objstream.Type = 1
	objstream.Mode =3
	objstream.Open
	objstream.Write body
	objstream.Position = 0
	objstream.Type = 2
	objstream.Charset = Cset
	BytesToBstr = objstream.ReadText
	objstream.Close
	set objstream=nothing
End Function

End Class

%>