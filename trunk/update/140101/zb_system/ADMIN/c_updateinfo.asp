<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)&(sipo)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_updateinfo.asp
'// 开始时间:    2007-1-26
'// 最后修改:    
'// 备    注:    
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<%

Response.ExpiresAbsolute   =   Now()   -   1           
Response.Expires   =   0   
Response.CacheControl   =   "no-cache"

If Len(ZC_UPDATE_INFO_URL)>0 Then

	Dim strPingTime
	Dim strPingContent

	Dim b,h
	b=False
	h=Now
	Dim fso,f
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(BlogPath & "zb_users\CACHE\updateinfo.txt")=True Then
		Set f = fso.GetFile(BlogPath & "zb_users\CACHE\updateinfo.txt")

		strPingContent=LoadFromFile(BlogPath & "zb_users\CACHE\updateinfo.txt","utf-8")
		If DateDiff("h",f.DateLastModified,Now)>24 Then
			b=True
		Else
			h=f.DateLastModified
		End If
	Else
		b=True
	End If

	If IsEmpty(Request.QueryString("reload"))=False Then
		b=True
	End If


	If b=True Then

		Dim objPing
		Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP")

		objPing.open "GET",ZC_UPDATE_INFO_URL,False

		objPing.send

		If objPing.ReadyState=4 Then
			If objPing.Status=200 Then
				strPingContent=objPing.responseText
			End If
		End If

		Dim objRegExp
		Set objRegExp=New RegExp
		objRegExp.IgnoreCase =True
		objRegExp.Global=True
		objRegExp.Pattern="<script.*/*>|</script>|<[a-zA-Z][^>]*=['""]+javascript:\w+.*['""]+>|<\w+[^>]*\son\w+=.*[ /]*>"
		strPingContent= objRegExp.Replace(strPingContent,"")

		Set objPing = Nothing

		Call SaveToFile(BlogPath & "zb_users\CACHE\updateinfo.txt",strPingContent,"utf-8",False)

	End If
	'Response.AddHeader "Last-Modified",ParseDateForRFC822GMT(h)
	Response.Write strPingContent

End If
%>