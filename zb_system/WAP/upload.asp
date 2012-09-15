<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../admin/ueditor/asp/ASPIncludeFile.asp"-->
<!-- #include file="c_system_wap.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style type="text/css">body,form,input{margin:0;padding:0;}</style>
<title>Upload</title>
</head>
<body>
<%
ShowError_Custom="Call ShowError_WAP(id)"
Call System_Initialize
Call CheckReference("")
If Not CheckRights("FileUpload") Then Call ShowError(6)
dim upload,file,state,uploadPath,PostTime,isPicture
Randomize
PostTime=GetTime(Now())
Dim strUPLOADDIR
strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"\"&Year(GetTime(Now()))&"\"&Month(GetTime(Now()))
Select Case Request.QueryString("act")
	Case "upload"
		Call CreatDirectoryByCustomDirectory(strUPLOADDIR)
		Set upload=New UpLoadClass
		upload.AutoSave=2
		upload.Charset="UTF-8"
		upload.FileType=Replace(ZC_UPLOAD_FILETYPE,"|","/")
		upload.savepath=BlogPath & strUPLOADDIR &"\"
		upload.maxsize=1024*1024*1024
		upload.open
		tExt=LCASE(upload.form("edtFileLoad_Ext"))
		If tExt="png" Or tExt="jpg" Or tExt="gif" Or tExt="jpeg" Or tExt="bmp" Then isPicture=True
		Dim Path
		Path=Replace(BlogPath & strUPLOADDIR &"\" & upload.form("edtFileLoad_Name")	,"\","/")
		Dim s
		FileName=BlogHost & strUPLOADDIR &"\" & upload.form("edtFileLoad_Name")
		Dim bolStu
		If isPicture Then
			bolStu=upload.Save("edtFileLoad",0)
		Else
			bolStu=upload.Save("edtFileLoad",1)
		End If
		If bolStu Then
			Dim uf
			Set uf=New TUpLoadFile
			uf.AuthorID=BlogUser.ID
			uf.AutoName=False
			uf.IsManual=True
			uf.FileSize=upload.form("edtFileLoad_Size")
			uf.FileName=upload.form("edtFileLoad")
			uf.UpLoad
		End If
		
		Dim nURL,tExt
		
		nURL=Replace(BlogHost & strUPLOADDIR & "/"&upload.form("edtFileLoad"),"\","/")
		If isPicture Then
			strData="<script type=""text/javascript"" language=""javascript"">try{top.document.getElementsByName('txaContent')[0].value+='<img src="""&nURL&""" alt="""&upload.form("edtFileLoad_Name")&""" title="""&upload.form("edtFileLoad_Name")&""" width="""&upload.Form("edtFileLoad_width")&""" height="""&upload.Form("edtFileLoad_height")&"""/>'}catch(e){}"
		Else
			strData="<script type=""text/javascript"" language=""javascript"">try{top.document.getElementsByName('txaContent')[0].value+='<a href="""&nUrl&""" target=""_blank""/>'}catch(e){}"
		End If
		strData=strData &"</script>上传完成！<a href=""upload.asp"">继续上传</a>"
		
		response.write strData
		response.end
End Select

%>
<form id="form1" name="form1" enctype="multipart/form-data" method="post" action="upload.asp?act=upload">
上传：<input type="file" name="edtFileLoad" id="edtFileLoad" /><input type="submit" name="button" id="button" value="上传" />
</form>
</body>
</html>
