<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
OPTION EXPLICIT
Server.ScriptTimeOut=5000
%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->

<!-- #include file="UpLoadClass.asp" -->
<%


Dim f_Name,f_SaveName,f_Path,f_Size,f_Ext,f_Err,f_Save,f_Time
dim FileUpload , FormName , rndtime
rndtime = Year(Now())&Month(Now()) 
CreatDirectoryByCustomDirectory("upload/" & rndtime)	'创建目录
FormName = "Filedata"		'文件域名称
set FileUpload = New UpLoadClass
FileUpload.SavePath = "upload/" & rndtime & "/"		'上传文件目录
FileUpload.Charset="UTF-8"
FileUpload.Open()


	f_Err = FileUpload.Form(FormName & "_Err")

			
	IF f_Err = 0 Then
		
			f_Name = FileUpload.Form(FormName & "_Name")					'原文件名
			f_SaveName = FileUpload.Form(FormName)								'保存文件名
			f_Path = FileUpload.SavePath													'保存路径
			f_Size = FileUpload.Form(FormName & "_Size")					'文件大小
			f_Ext = FileUpload.Form(FormName & "_Ext")						'文件类型
			f_Time = Now()																	'保存时间
			
set FileUpload = nothing
			
'ASPJPEG处理

dim FilePath,FileName,FileNamelen,f_SaveName1,f_s_SaveName,imgWidth,imgHeight



FilePath = Server.MapPath("./")& "\upload\" & rndtime   '设置上传目录位置
FileName = FilePath&"\"&f_SaveName

Dim Jpeg
Set Jpeg = Server.CreateObject("Persits.Jpeg")

'如果aspjpeg版本大于1.9，启用保护Metadata
If Jpeg.Version>= "1.9" then Jpeg.PreserveMetadata = True

Jpeg.Open(FileName)

'变更缩略图文件扩展名为jpg
FileNamelen = Len(f_SaveName) - 4
f_SaveName1 = f_SaveName
f_s_SaveName = "small_"&Left(f_SaveName, FileNamelen) &".jpg"

'缩略图处理，判断哪边为长边，以长边进行缩放
imgWidth = Jpeg.OriginalWidth
imgHeight = Jpeg.OriginalHeight

If imgWidth>= imgHeight And imgWidth>45 Then
	Jpeg.Width = 45
	Jpeg.Height = Jpeg.OriginalHeight / (Jpeg.OriginalWidth / 45)
End If
If imgHeight>imgWidth And imgHeight>60 Then
	Jpeg.Height = 60
	Jpeg.Width = Jpeg.OriginalWidth / (Jpeg.OriginalHeight / 60)
End If

'保存缩略图，并进行微度锐化
Jpeg.Sharpen 1, 110
Jpeg.Save (FilePath & "\" & f_s_SaveName)
Jpeg.Close:Set Jpeg = Nothing 
'ASPJPEG处理结束
			
			Response.Write("{""name"":""" & f_Name & """,""savename"":""" & rndtime & "/" & f_SaveName & """,""s_savename"":""" & rndtime & "/" & f_s_SaveName & """,""path"":""" & f_Path & """,""size"":" & f_Size & ",""ext"":""" & f_Ext & """,""time"":""" & f_Time & """,""err"":" & f_Err & "}")
		
		Else
			
			Response.Write("{""err"":" & f_Err & "}")
			
	End IF
			
			

%>