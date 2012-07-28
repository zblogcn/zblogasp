<%@ CODEPAGE=65001 %>
<!--#include file="up_inc.asp"-->
<!-- #include file="..\..\..\..\zb_users\c_option.asp" -->
<!-- #include file="..\..\..\function\c_function.asp" -->
<!-- #include file="..\..\..\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\function\c_system_base.asp" -->
<!-- #include file="..\..\..\function\c_system_event.asp" -->
<!-- #include file="..\..\..\function\c_system_plugin.asp" -->
<!-- #include file="..\..\..\..\zb_users\plugin\p_config.asp" -->
<%
dim upload,file,state,uploadPath,PostTime
Randomize
Call System_Initialize
PostTime=GetTime(Now())
Dim strUPLOADDIR

strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"\"&Year(GetTime(Now()))&"\"&Month(GetTime(Now()))
CreatDirectoryByCustomDirectory(strUPLOADDIR)

Set upload=New UpLoadClass
upload.AutoSave=2
upload.Charset="UTF-8"
upload.FileType=Replace(ZC_UPLOAD_FILETYPE,"|","/")
upload.savepath=BlogPath &  strUPLOADDIR &"\"
upload.maxsize=1024*1024*1024
upload.open

Set BlogUser=Nothing
Set BlogUser =New TUser
BlogUser.LoginType="Self"
BlogUser.name=CStr(Trim(upload.form("username")))
BlogUser.Password=CStr(Trim(upload.form("password")))
BlogUser.Verify()



If Not CheckRights("ArticleEdt") Then Call ShowError(6)
For Each sAction_Plugin_FileUpload_Begin in Action_Plugin_FileUpload_Begin
	If Not IsEmpty(sAction_Plugin_FileUpload_Begin) Then Call Execute(sAction_Plugin_FileUpload_Begin)
Next


Dim Path
Path=Replace(BlogPath &  strUPLOADDIR &"\" & upload.form("edtFileLoad_Name")	,"\","/")
Dim s
FileName=GetCurrentHost& strUPLOADDIR &"\" & upload.form("edtFileLoad_Name")
s=upload.Save("edtFileLoad",1)
objConn.Execute("INSERT INTO [blog_UpLoad]([ul_AuthorID],[ul_FileSize],[ul_FileName],[ul_PostTime],[ul_FileIntro],[ul_DirByTime]) VALUES ("& BlogUser.ID &",'"& upload.form("edtFileLoad_Size") &"','"& upload.form("edtFileLoad") &"','"& PostTime &"','Attatment',"&CInt(ZC_UPLOAD_DIRBYMONTH)&")")

response.Write "{'state':'"& upload.Error2Info("edtFileLoad") & "','url':'"& upload.form("edtFileLoad_Name") &"','fileType':'"&upload.form("edtFileLoad_Ext")&"','original':'"& upload.form("edtFileLoad_Name")&"'}"

	
For Each sAction_Plugin_uEditor_FileUpload_End in Action_Plugin_uEditor_FileUpload_End
	If Not IsEmpty(sAction_Plugin_uEditor_FileUpload_End) Then Call Execute(sAction_Plugin_uEditor_FileUpload_End)
Next

set upload=nothing
Call System_Terminate()

%>