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
				Response.Write "try{eval(""$(\""#"& aryElement(0) &"\"").html(\"""& LoadCountInfo(aryElement(1)) &"\"")"");}catch(e){}"
			Next
			'LoadSidebar
		End If
		s=Request.QueryString("inculde")
		If s<>"" Then
			strPara=s
			aryPara=Split(strPara,",")

			For l=0 To UBound(aryPara)-1
				aryElement=Split(aryPara(l),"=")
				Response.Write "try {eval(""{$('#" & aryElement(0) & "').replaceWith('" & LoadFileInfo(aryElement(1)) & "')}"");} catch (e) {}"
			Next
			'LoadSidebar
			
		End If

		s=Request.QueryString("count")
		If s<>"" Then
			strPara=s
			aryPara=Split(strPara,",")

			For l=0 To UBound(aryPara)-1
				aryElement=Split(aryPara(l),"=")
				Response.Write "try{eval(""$(\""#"& aryElement(0) &"\"").html(\"""& CStr(LoadCountInfo(aryElement(1))+1) &"\"")"");}catch(e){}"
				Call UpdateCountInfo(aryElement(1))
			Next
		
		End If
		LoadSidebar

	ElseIf f="autoinfo" Then

		Call OpenConnect()

		Set BlogUser = New TUser
		If BlogUser.Verify()=True Then
			Response.Write "try{$('#inpName').val('"&BlogUser.Name&"');}catch(e){}"
			Response.Write "try{$('#inpEmail').val('"&BlogUser.Email&"');}catch(e){}"
			Response.Write "try{$('#inpHomePage').val('"&BlogUser.HomePage&"');}catch(e){}"
			Response.Write "try{$('.cp-hello').html('"&Replace(ZC_MSG023,"%s",BlogUser.FirstName) & " (" & ZVA_User_Level_Name(BlogUser.Level)&")');"
			Response.Write "$('.cp-login').find('a').html('["&ZC_MSG248&"]');"
			If CheckRights("ArticleEdt")=True Then
				Response.Write "$('.cp-vrs').find('a').html('["&ZC_MSG168&"]');$('.cp-vrs').find('a').attr('href','"&BlogHost&"zb_system/cmd.asp?act=ArticleEdt');"
			End IF
			Response.Write "}catch(e){}"
		End If
		
		Call CloseConnect()

	End If 
	
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

	If f="now" Then f=Year(Date)&"-"&Month(Date)

	s=Replace(MakeCalendar(f),"<#ZC_BLOG_HOST#>",BlogHost)

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

	Dim aryArticleCount()
	Redim Preserve aryArticleCount(objDS(0,UBound(objDS, 2)))
	
	For i = 0 To UBound(objDS, 2)
		aryArticleCount(objDS(0,i))=objDS(1,i)
	Next

	Application.Lock
	Application(ZC_BLOG_CLSID&"CACHE_ARTICLE_VIEWCOUNT")=aryArticleCount
	Application.UnLock

	ReadCountInfo=aryArticleCount

End Function
'*********************************************************




'*********************************************************
' 目的：    
' 输入：    
' 返回：    
'*********************************************************
Function UpdateCountInfo(id)

	Call CheckParameter(id,"int",0)

	Call OpenConnect()

	objConn.Execute("UPDATE [blog_Article] SET [log_ViewNums]=[log_ViewNums]+1 WHERE [log_ID] =" & id)

	Call CloseConnect()

	Dim aryArticleCount
	Application.Lock
	aryArticleCount=Application(ZC_BLOG_CLSID&"CACHE_ARTICLE_VIEWCOUNT")
	aryArticleCount(id)=aryArticleCount(id)+1
	Application(ZC_BLOG_CLSID&"CACHE_ARTICLE_VIEWCOUNT")=aryArticleCount
	Application.UnLock

End Function
'*********************************************************




'*********************************************************
' 目的：    
' 输入：    
' 返回：    
'*********************************************************
Function LoadCountInfo(id)

	Dim aryArticleCount

	Application.Lock
	aryArticleCount=Application(ZC_BLOG_CLSID&"CACHE_ARTICLE_VIEWCOUNT")
	Application.UnLock

	If IsEmpty(aryArticleCount) Then
		aryArticleCount=ReadCountInfo
	End If

	LoadCountInfo=aryArticleCount(id)

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

	For i=0 To UBound(aryTemplateTagsName)
		If aryTemplateTagsName(i)="ZC_BLOG_HOST" Then
			aryTemplateTagsValue(i)=BlogHost
		End If 
	Next

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
	strContent= Replace(strContent,"""","\""")
	strContent= Replace(strContent,vbCrLf,"")
	strContent= Replace(strContent,vbLf,"")

	LoadFileInfo=strContent

End Function
'*********************************************************

Dim isSidebarLoad
isSidebarLoad=False
Function LoadSidebar()
	If isSidebarLoad Then Exit Function
	Response.Write "try{sidebarloaded.execute()}catch(e){}"
	LoadSidebar=True
	isSidebarLoad=True
End Function
%>