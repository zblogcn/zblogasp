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

Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

Dim objUpload
Set objUpload=New UpLoadClass
objUpload.AutoSave=2
objUpload.Charset="utf-8"
objUpload.FileType="jpg/png/bmp/gif"
objUpload.savepath=BlogPath & "zb_users\theme\default\include\"
objUpload.maxsize=ZC_UPLOAD_FILESIZE
objUpload.open


Dim m,s
For Each s In objUpload.FileItem
	Select Case s
	Case "include_bg-nav.jpg"
		If objUpload.Form("include_bg-nav.jpg_Width")=1600 And objUpload.Form("include_bg-nav.jpg_Height")=180 Then
			objUpload.Save "include_bg-nav.jpg","bg-nav.jpg"
			If objUpload.Form("include_bg-nav.jpg_Err")=0 Then
				ClearGlobeCache
				LoadGlobeCache
				BlogRebuild_Default
				Call SetBlogHint(True,Empty,Empty)
			Else
				Call SetBlogHint_Custom(objUpload.Error2Info("include_bg-nav.jpg"))
			End If
		Else
			Call SetBlogHint_Custom("分辨率必须符合1600x180!")
		End If
	End Select
Next

Response.Redirect "editor.asp"
%>

<%Call System_Terminate()%>
