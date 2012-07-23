<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_html_js.asp
'// 开始时间:    2005.02.22
'// 最后修改:    
'// 备    注:    html模板脚本辅助
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<% Response.ContentType="application/x-javascript" %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="../function/c_function.asp" -->
<!-- #include file="../function/c_system_lib.asp" -->
<!-- #include file="../function/c_system_base.asp" -->
<!-- #include file="../function/c_system_event.asp" -->
<!-- #include file="../function/c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->
<%
'/////////////////////////////////////////////////////////////////////////////////////////

Call LoadGlobeCache

Dim s
Dim f

f=Request.QueryString("act")
If f<>"" Then

	If (f="ArticleView") Or (f="count") Then

		Dim i,j

		i=Request.QueryString("id")
		Call CheckParameter(i,"int",0)

		If i=0 Then Response.End

		Response.Clear
		Response.Write "document.write(""" & CStr(LoadCountInfo(i)+1) & """)"

		Call UpdateCountInfo(i)

		Response.End


	ElseIf f="batch" Then

		Dim strPara,aryPara,aryElement
		Dim k,l

		s=Request.QueryString("view")
		If s<>"" Then
			strPara=s
			aryPara=Split(strPara,",")

			For l=0 To UBound(aryPara)-1
				aryElement=Split(aryPara(l),"=")
				Response.Write "try{eval(""document.getElementById(\"""& aryElement(0) &"\"").innerHTML=\"""& LoadCountInfo(aryElement(1)) &"\"""");}catch(e){}"
			Next

		End If
		s=Request.QueryString("inculde")
		If s<>"" Then
			strPara=s
			aryPara=Split(strPara,",")

			For l=0 To UBound(aryPara)-1
				aryElement=Split(aryPara(l),"=")
				Response.Write "try{eval(""  {$(\""#"& aryElement(0) &"\"").after(\"""& LoadFileInfo(aryElement(1)) &"\"");$(\""#"& aryElement(0) &"\"").remove();}  "");}catch(e){}"
			Next

		End If

		s=Request.QueryString("count")
		If s<>"" Then
			strPara=s
			aryPara=Split(strPara,",")

			For l=0 To UBound(aryPara)-1
				aryElement=Split(aryPara(l),"=")
				Response.Write "try{eval(""document.getElementById(\"""& aryElement(0) &"\"").innerHTML=\"""& CStr(LoadCountInfo(aryElement(1))+1) &"\"""");}catch(e){}"
				Call UpdateCountInfo(aryElement(1))
			Next

		End If
	End If

	'日历当天高亮显示
	Response.Write "try{eval(""document.getElementById(\""pCalendar_\""+(new Date()).getFullYear()+\""_\""+((new Date()).getMonth()+1)+\""_\""+(new Date()).getDate()).className+=\"" cd\"""");}catch(e){};"

	Response.End

End If


'/////////////////////////////////////////////////////////////////////////////////////////


f=Request.QueryString("include")
If f<>"" Then

	Response.Clear
	Response.Write "document.write(""" & LoadFileInfo(f) & """)"
	Response.End

End If


'/////////////////////////////////////////////////////////////////////////////////////////


f=Request.QueryString("date")
If f<>"" Then

	Call System_Initialize()

	f=Request.QueryString("date")

	s=MakeCalendar(f)

	Response.Clear
	Response.Write "document.write(""" & Replace(s,"""","\""") & """)"
	Response.End

	Call System_Terminate()

End If
'/////////////////////////////////////////////////////////////////////////////////////////




'*********************************************************
' 目的：    
' 输入：    
' 返回：    
'*********************************************************
Function ReadCountInfo()

	Call OpenConnect()

	Dim objRS,i,j,objDS
	Set objRS=objConn.Execute("SELECT [log_ID],[log_ViewNums] FROM [blog_Article] ORDER BY [log_ID] ASC")
		If (not objRS.bof) And (not objRS.eof) Then
			objDS=objRS.GetRows
		End If
		objRS.Close
	Set objRS=Nothing

	Call CloseConnect()

	If IsNull(objDS) or IsEmpty(objDS) Then ReadCountInfo=Empty : Exit Function

	Dim aryCount()
	Redim Preserve aryCount(objDS(0,UBound(objDS, 2)))
	
	For i = 0 To UBound(objDS, 2)
		aryCount(objDS(0,i))=objDS(1,i)
	Next

	Application.Lock
	Application(ZC_BLOG_CLSID&"CACHE_ARTICLE_VIEWCOUNT")=aryCount
	Application.UnLock

	ReadCountInfo=aryCount

End Function
'*********************************************************




'*********************************************************
' 目的：    
' 输入：    
' 返回：    
'*********************************************************
Function UpdateCountInfo(id)

	Call CheckParameter(id,"int",0)

	Dim strLastUpdate,intCount,aryCount,objRS,bolToDb
	bolToDb=False

	Application.Lock
	strLastUpdate=Application(ZC_BLOG_CLSID&"LAST_UPDATE")
	aryCount=Application(ZC_BLOG_CLSID&"CACHE_ARTICLE_VIEWCOUNT")
	aryCount(id)=aryCount(id)+1
	Application(ZC_BLOG_CLSID&"CACHE_ARTICLE_VIEWCOUNT")=aryCount
	Application.UnLock

	If IsEmpty(strLastUpdate) Or Not IsDate(strLastUpdate) Then
		Application.Lock
		Application(ZC_BLOG_CLSID&"LAST_UPDATE") = Now()
		strLastUpdate = Application(ZC_BLOG_CLSID&"LAST_UPDATE")
		Application.UnLock
		bolToDb=True
	End If
	If DateDiff("s",strLastUpdate,Now()) > 60 Then bolToDb=True
 	If bolToDb=True Then
 		Call OpenConnect()
		Set objRS=objConn.Execute("SELECT [log_ViewNums] FROM [blog_Article] WHERE [log_ID] =" & id)
		If (not objRS.bof) And (not objRS.eof) Then
			intCount=objRS("log_ViewNums")
		Else
			intCount=0
		End If
  		Set objRS=Nothing
		If aryCount(id) > intCount Then
			objConn.Execute("UPDATE [blog_Article] SET [log_ViewNums]=" & CLng(aryCount(id)) & " WHERE [log_ID] =" & id)
			Application.Lock
			Application(ZC_BLOG_CLSID&"LAST_UPDATE") = Now()
			Application.UnLock
		Else
			aryCount(id) = intCount
		  	Application.Lock
		  	Application(ZC_BLOG_CLSID&"CACHE_ARTICLE_VIEWCOUNT")=aryCount
		  	Application.UnLock
  		End If
		Call CloseConnect()
 	End If
End Function
'*********************************************************




'*********************************************************
' 目的：    
' 输入：    
' 返回：    
'*********************************************************
Function LoadCountInfo(id)

	Dim aryCount

	Application.Lock
	aryCount=Application(ZC_BLOG_CLSID&"CACHE_ARTICLE_VIEWCOUNT")
	Application.UnLock

	If IsEmpty(aryCount) Then
		aryCount=ReadCountInfo
	End If

	LoadCountInfo=aryCount(id)

End Function
'*********************************************************




'*********************************************************
' 目的：    
' 输入：    
' 返回：    
'*********************************************************
Function LoadFileInfo(name)

	Dim strContent
	Dim objStream

	Dim i,j

	Dim aryTemplateTagsName
	Dim aryTemplateTagsValue

	Application.Lock
	aryTemplateTagsName=Application(ZC_BLOG_CLSID & "TemplateTagsName")
	aryTemplateTagsValue=Application(ZC_BLOG_CLSID & "TemplateTagsValue")
	Application.UnLock

	j=UBound(aryTemplateTagsName)

	For i=1 to j
		If aryTemplateTagsName(i)="TEMPLATE_INCLUDE_" & UCase(name) Then
			strContent=aryTemplateTagsValue(i)
			Exit For
		ElseIf aryTemplateTagsName(i)="CACHE_INCLUDE_" & UCase(name) Then
			strContent=aryTemplateTagsValue(i)
		End If
	Next

	j=UBound(aryTemplateTagsName)

	For i=1 to j
		strContent=Replace(strContent,"<#" & aryTemplateTagsName(i) & "#>",aryTemplateTagsValue(i))
	Next

	strContent= Replace(strContent,"\","\\")
	strContent= Replace(strContent,"/","\/")
	strContent= Replace(strContent,"""","'")
	strContent= Replace(strContent,vbCrLf,"")
	strContent= Replace(strContent,vbLf,"")

	LoadFileInfo=strContent

End Function
'*********************************************************
%>