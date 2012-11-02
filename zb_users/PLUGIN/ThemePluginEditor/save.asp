<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ThemePluginEditor")=False Then Call ShowError(48)

Dim aryReq(),tmpary(),t,s,i
i=0
For Each s In Request.Form
	If Left(s,8)="include_" Then
		Redim Preserve aryReq(i)
		Redim tmpary(2)
		tmpary(0)=Request.Form(s)
		tmpary(1)=Request.Form("type_"& Right(s,Len(s)-8))
		tmpary(2)=Right(s,Len(s)-8)
		aryReq(i)=tmpary
		i=i+1
	End If
Next




Dim aryName(),aryDesc(),aryTr(),strTemplateName,strTr
strTr=LoadFromFile(BlogPath & "\zb_users\plugin\themeplugineditor\resources\tr_template.asp","utf-8")
Redim aryName(i-1)
Redim aryDesc(i-1)
Redim aryTr(i-1)
Dim arySave,j
arySave=Array(Array(),Array(),Array(),Array())
arySave(0)=Array("save.asp",LoadFromFile(BlogPath & "\zb_users\plugin\themeplugineditor\resources\save.asp","utf-8"))
arySave(1)=Array("editor.asp",LoadFromFile(BlogPath & "\zb_users\plugin\themeplugineditor\resources\editor.asp","utf-8"))
arySave(2)=Array("tr_template.asp",strTr)
arySave(3)=Array("include.asp",LoadFromFile(BlogPath & "\zb_users\plugin\themeplugineditor\resources\include.asp","utf-8"))
strTemplateName=ZC_BLOG_THEME
For s=0 To Ubound(aryReq)
	aryTr(s)=strTr
	aryTr(s)=Replace(aryTr(s),"<%=主题名%"&">",strTemplateName)
	aryTr(s)=Replace(aryTr(s),"<%=文件注释%"&">",aryReq(s)(0))
	aryTr(s)=Replace(aryTr(s),"<%=文件名%"&">",aryReq(s)(2))		
Next
arySave(2)(1)=Join(aryTr,vbCrlf)
For s=0 To 3
	If s<>2 Then
		arySave(s)(1)=Replace(arySave(s)(1),"<%=主题名%"&">",strTemplateName)
		arySave(s)(1)=Replace(arySave(s)(1),"<%=表格%"&">",arySave(2)(1))
		arySave(s)(1)=Replace(arySave(s)(1),"<%=文件注释数组%"&">",""""&Join2(aryReq,0,""",""")&"""")
		arySave(s)(1)=Replace(arySave(s)(1),"<%=文件名数组%"&">",""""&Join2(aryReq,1,",")&"""")
		Call SaveToFile(BlogPath & "\zb_users\theme\" & strTemplateName & "\plugin\" & arySave(s)(0),arySave(s)(1),"utf-8",False)
	End If
	
Next

'Stop

Function Join2(ary,int,s)
	Dim i,ary2()
	Redim ary2(Ubound(ary))
	For i=0 To Ubound(ary)
		ary2(i)=ary(i)(int)
	Next
	Join2=Join(ary2,s)
End Function
%>
