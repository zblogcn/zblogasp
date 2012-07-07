<!--#include file="up_inc.asp"-->
<!-- #include file="..\..\..\..\zb_users\c_option.asp" -->
<!-- #include file="..\..\..\function\c_function.asp" -->
<!-- #include file="..\..\..\function\c_function_md5.asp" -->
<!-- #include file="..\..\..\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\function\c_system_base.asp" -->
<!-- #include file="..\..\..\function\c_system_event.asp" -->
<!-- #include file="..\..\..\function\c_system_plugin.asp" -->
<!-- #include file="..\..\..\function\rss_lib.asp" -->
<!-- #include file="..\..\..\function\atom_lib.asp" -->
<!-- #include file="..\..\..\..\zb_users\plugin\p_config.asp" -->
<%
'On Error Resume Next
Call System_Initialize()
Call CheckReference("")
If Not CheckRights("ArticleEdt") Then Call ShowError(6)
For Each sAction_Plugin_FileUpload_Begin in Action_Plugin_FileUpload_Begin
	If Not IsEmpty(sAction_Plugin_FileUpload_Begin) Then Call Execute(sAction_Plugin_FileUpload_Begin)
Next


dim upload,file,formName,title,state,picSize,picType,uploadPath,fileExt,fileName,prefix,PostTime
picSize = 200                        '允许的文件大小，单位KB
state="SUCCESS" 
set upload=new upload_5xSoft         '建立上传对象
title = htmlspecialchars(upload.form("pictitle"))
for each formName in upload.file
	Randomize
	set file=upload.file(formName)
	if file.filesize > picSize*1024 then
   	        state="文件大小超出限制"
    end if
    fileExt = split(lcase(mid(file.FileName, instrrev(file.FileName,"."))),".")(1)
	
    if instr(ZC_UPLOAD_FILETYPE, fileExt)=0 then
         state = "文件类型错误"
    end If
	PostTime=GetTime(Now())
	Dim strUPLOADDIR
	If ZC_UPLOAD_DIRBYMONTH Then
			strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"\"&Year(GetTime(Now()))&"\"&Month(GetTime(Now()))
			CreatDirectoryByCustomDirectory("zb_users\"&strUPLOADDIR)
	Else
			strUPLOADDIR = ZC_UPLOAD_DIRECTORY
	End If
	If Fileext="jpg" or Fileext="jpeg" or Fileext="png" or Fileext="gif" or Fileext="bmp" or Fileext="ico" or Fileext="tiff" then 	file.FileName=int(Rnd*100000000000000)&"."&Fileext
	Dim Path
	Path=Replace(BlogPath & "zb_users\"& strUPLOADDIR &"\" & file.FileName,"\","/")
	If state="SUCCESS" then
		file.SaveAs Path
    end if
	If err.number<>0 Then state=err.description
	objConn.Execute("INSERT INTO [blog_UpLoad]([ul_AuthorID],[ul_FileSize],[ul_FileName],[ul_PostTime],[ul_FileIntro],[ul_DirByTime]) VALUES ("& BlogUser.ID &","& file.filesize &",'"& file.FileName &"','"& PostTime &"','PicOrAttatment',"&CInt(ZC_UPLOAD_DIRBYMONTH)&")")

	response.Write "{'state':'"& state & "','url':'"& file.FileName &"','fileType':'"&fileExt&"'}"
    set file=nothing
	
	
next
For Each sAction_Plugin_uEditor_FileUpload_End in Action_Plugin_uEditor_FileUpload_End
	If Not IsEmpty(sAction_Plugin_uEditor_FileUpload_End) Then Call Execute(sAction_Plugin_uEditor_FileUpload_End)
Next

set upload=nothing
Call System_Terminate()
    function htmlspecialchars(someString)
        htmlspecialchars = replace(replace(replace(replace(someString, "&", "&amp;"), ">", "&gt;"), "<", "&lt;"), """", "&quot;")
    end function
%>