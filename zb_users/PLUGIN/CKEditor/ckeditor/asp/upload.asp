<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% Response.CodePage=65001 %>
<% Response.Charset="UTF-8" %>
<!--#include file="UpLoad_Class.asp"-->
<!--#include file="ASPIncludeFile.asp"-->
<%
' CKEditor ASP
'未寒
CKEupload

Dim aspUrl, savePath, saveUrl, maxSize, fileName, fileExt, newFileName, filePath, fileUrl, dirName
Dim extStr, imageExtStr, flashExtStr, mediaExtStr, fileExtStr
Dim upload, file, fso, ranNum, hash, ymd, mm, dd, result, CKEditorFuncNum

aspUrl = Request.ServerVariables("SCRIPT_NAME")
aspUrl = left(aspUrl, InStrRev(aspUrl, "/"))

'文件保存目录路径
savePath = "../../../../../" & ZC_UPLOAD_DIRECTORY & "/"
'文件保存目录URL
saveUrl = ZC_BLOG_HOST & ZC_UPLOAD_DIRECTORY & "/"
'定义允许上传的文件扩展名

fileExtStr = ZC_UPLOAD_FILETYPE
'最大文件大小
maxSize = ZC_UPLOAD_FILESIZE '5 * 1024 * 1024 '5M

Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(Server.mappath(savePath)) Then
	showError(saveUrl&"上传目录不存在。"&savePath)
End If

CKEditorFuncNum = Request.QueryString("CKEditorFuncNum")

If instr(lcase("image,flash,media,file"), dirName) < 1 Then
	showError("目录名不正确。")
End If


set upload = new AnUpLoad
upload.Exe = fileExtStr
upload.MaxSize = maxSize
upload.GetData()
if upload.ErrorID>0 then 
	showError(upload.Description)
end if

mm = month(now)
savePath = savePath & year(now) & "/" & mm & "/"
saveUrl = saveUrl & year(now) & "/" & mm & "/"

If Not fso.FolderExists(Server.mappath(savePath)) Then
	fso.CreateFolder(Server.mappath(savePath))
End If

set file = upload.files("upload")
if file is nothing then
	showError("请选择文件。")
end if

set result = file.saveToFile(savePath, 0, true)
if result.error then
	showError(file.Exception)
end if

filePath = Server.mappath(savePath & file.filename)
fileUrl = saveUrl & file.filename
Dim filenameupload,filesizeupload

filenameupload=file.filename
filesizeupload=upload.TotalSize

Set upload = nothing
Set file = nothing

If Not fso.FileExists(filePath) Then
	showError("上传文件失败。")
End If

	Dim uf
	Set uf=New TUpLoadFile
	uf.AuthorID=BlogUser.ID
	uf.AutoName=False
	uf.IsManual=True
	uf.FileSize= filesizeupload
	uf.FileName= filenameupload
	uf.UpLoad

Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
response.write("<script type='text/javascript'>")
response.write "window.parent.CKEDITOR.tools.callFunction(" & CKEditorFuncNum & ",'" & fileUrl & "','')"
response.write("</script>")
Response.End

Function showError(message)
	Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
	response.write("<script type='text/javascript'>")
	response.write "window.parent.CKEDITOR.tools.callFunction(" & CKEditorFuncNum & ",'','" & message & "')"
	response.write("</script>")
	Response.End
End Function
%>
