<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'************************************
' Powered by ThemePluginEditor 1.1
' zsx http://www.zsxsoft.com
'************************************
%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\..\..\zb_system\admin\ueditor\asp\aspincludefile.asp" -->
<%
Response.Write "{""succeed"":"
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
objUpload.savepath=BlogPath & "zb_users\theme\metro\style\images\"
objUpload.maxsize=ZC_UPLOAD_FILESIZE
objUpload.open


Dim m,s
For Each s In objUpload.FileItem
	'If Left(s,7)="include" Then
		m=objUpload.Save(s,s)
	'End If
	Response.Write "0"
'	Response.Write """"&objUpload.Error2Info(s)&""""
Next
Response.Write "}"
%>

<%Call System_Terminate()%>
