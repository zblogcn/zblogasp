<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../../zb_system/admin/ueditor/asp/aspincludefile.asp"-->
<!-- #include file="function.asp"-->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

If CheckPluginState("AppCentre")=False Then Call ShowError(48)




Dim objUpLoadFile
Set objUpLoadFile=New TUpLoadFile

Dim objUpload,isOK
Set objUpload=New UpLoadClass
objUpload.AutoSave=2
objUpload.Charset="utf-8"
objUpload.FileType="zba"
objUpload.savepath=BlogPath & "zb_users\cache\"
objUpload.maxsize=ZC_UPLOAD_FILESIZE
objUpload.open
If objUpload.Save("edtFileLoad",0)=True Then
	Call InstallApp(BlogPath & "zb_users\cache\"&objUpload.form("edtFileLoad"))
	CreateObject("scripting.filesystemobject").DeleteFile BlogPath & "zb_users\cache\*.zba"
Else
	If objUpload.Form("edtFileLoad_Ext")<>"zba" Then
		SetBlogHint_Custom "该应用不是Z-Blog 2.0应用，无法应用于Z-Blog 2.0！"
	Else
		SetBlogHint_Custom objUpload.Error2Info("edtFileLoad")
	End If 
End If


Response.Redirect Request.ServerVariables("HTTP_REFERER")


%>