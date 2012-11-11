<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'************************************
' Powered by ThemePluginEditor<%=版本号%>
' zsx http://www.zsxsoft.com
'************************************
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\..\..\zb_system\admin\ueditor\asp\aspincludefile.asp" -->

<%

Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

Dim objUpload
Set objUpload=New UpLoadClass
objUpload.AutoSave=2
objUpload.Charset="utf-8"
objUpload.FileType=Replace(ZC_UPLOAD_FILETYPE,"|","/")
objUpload.savepath=BlogPath & "zb_users\theme\<%=主题名%>\include\"
objUpload.maxsize=ZC_UPLOAD_FILESIZE
objUpload.open


Dim s
For Each s In objUpload.FormItem
	If Left(s,7)="include" Then
		Call SaveToFile(BlogPath & "zb_users/theme/<%=主题名%>/include/" & Right(s,Len(s)-8),objUpload.Form(s),"utf-8",False)
	End If
Next
Dim m
For Each s In objUpload.FileItem
	If Left(s,7)="include" Then
		m=objUpload.Save(s,Right(s,Len(s)-8))
	End If
Next
ClearGlobeCache
LoadGlobeCache
BlogRebuild_Default
Call SetBlogHint(True,Empty,Empty)
Response.Redirect "editor.asp"
%>

<%Call System_Terminate()%>
