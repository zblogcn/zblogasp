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
Dim PluginVer
PluginVer=" 1.0"
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("ThemePluginEditor")=False Then Call ShowError(48)
Dim HaveUpload
HaveUpload=False
Dim aryReq(),tmpary(),t,s,i
i=0
For Each s In Request.Form
	If Left(s,8)="include_" Then
		Redim Preserve aryReq(i)
		Redim tmpary(2)
		tmpary(0)=Request.Form(s)
		tmpary(1)=Request.Form("type_"& Right(s,Len(s)-8))
		If tmpary(1)=2 Then HaveUpload=True
		tmpary(2)=Right(s,Len(s)-8)
		aryReq(i)=tmpary
		i=i+1
	ElseIf Left(s,4)="new_" Then
		Call SaveToFile(BlogPath & "\zb_users\theme\" & ZC_BLOG_THEME & "\include\" & Right(s,Len(s)-4),"","utf-8",False)
	End If
	
Next




Dim aryName(),aryDesc(),aryTr(),strTemplateName,strTr,strUpload
strTr=LoadFromFile(BlogPath & "\zb_users\plugin\themeplugineditor\resources\tr_template.asp","utf-8")
strUpload=LoadFromFile(BlogPath & "\zb_users\plugin\themeplugineditor\resources\tr_template_upload.asp","utf-8")
Redim aryName(i-1)
Redim aryDesc(i-1)
Redim aryTr(i-1)
Dim arySave,j
arySave=Array(Array(),Array(),Array(),Array())
arySave(0)=Array("save.asp",LoadFromFile(BlogPath & "\zb_users\plugin\themeplugineditor\resources\"&IIf(HaveUpload,"save_upload.asp","save.asp"),"utf-8"))
arySave(1)=Array("editor.asp",LoadFromFile(BlogPath & "\zb_users\plugin\themeplugineditor\resources\"&IIf(HaveUpload,"editor_upload.asp","editor.asp"),"utf-8"))
arySave(2)=Array("tr_template.asp","")
arySave(3)=Array("include.asp",LoadFromFile(BlogPath & "\zb_users\plugin\themeplugineditor\resources\include.asp","utf-8"))
strTemplateName=ZC_BLOG_THEME
For s=0 To Ubound(aryReq)
	aryTr(s)=IIf(aryReq(s)(1)=2,strUpload,strTr)
	aryTr(s)=Replace(aryTr(s),"<%=主题名%"&">",strTemplateName)
	aryTr(s)=Replace(aryTr(s),"<%=文件注释%"&">",aryReq(s)(0))
	aryTr(s)=Replace(aryTr(s),"<%=文件名%"&">",aryReq(s)(2))	
	aryTr(s)=Replace(aryTr(s),"<%=主题调用代码%"&">","&lt;#TEMPLATE_INCLUDE_"&UCase(Split(aryReq(s)(2),".")(0))&"#&gt;")
Next
arySave(2)(1)=Join(aryTr,vbCrlf)
For s=0 To 3
	If s<>2 Then
		arySave(s)(1)=Replace(arySave(s)(1),"<%=主题名%"&">",strTemplateName)
		arySave(s)(1)=Replace(arySave(s)(1),"<%=表格%"&">",arySave(2)(1))
		arySave(s)(1)=Replace(arySave(s)(1),"<%=文件注释数组%"&">",""""&Join2(aryReq,0,""",""")&"""")
		arySave(s)(1)=Replace(arySave(s)(1),"<%=文件名数组%"&">",""""&Join2(aryReq,2,""",""")&"""")
		arySave(s)(1)=Replace(arySave(s)(1),"<%=版本号%"&">",PluginVer)
		Call SaveToFile(BlogPath & "\zb_users\theme\" & strTemplateName & "\plugin\" & arySave(s)(0),arySave(s)(1),"utf-8",False)
	End If
	
Next
Call SetBlogHint(True,Empty,Empty)
Response.Redirect "howtouse.asp"
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
