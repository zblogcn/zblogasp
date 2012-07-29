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
Call System_Initialize()
Call CheckReference("")
If Not CheckRights("FileUpload") Then Call ShowError(6)
For Each sAction_Plugin_FileUpload_Begin in Action_Plugin_FileUpload_Begin
	If Not IsEmpty(sAction_Plugin_FileUpload_Begin) Then Call Execute(sAction_Plugin_FileUpload_Begin)
Next

dim upload,file,state,uploadPath,PostTime
Randomize

PostTime=GetTime(Now())
Dim strUPLOADDIR

strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"\"&Year(GetTime(Now()))&"\"&Month(GetTime(Now()))
CreatDirectoryByCustomDirectory(strUPLOADDIR)

Set upload=New UpLoadClass
upload.AutoSave=2
upload.Charset="UTF-8"
upload.FileType=Replace(ZC_UPLOAD_FILETYPE,"|","/")
upload.savepath=BlogPath & strUPLOADDIR &"\"
upload.maxsize=1024*1024*1024
upload.open
Dim Path
Path=Replace(BlogPath & strUPLOADDIR &"\" & upload.form("edtFileLoad_Name")	,"\","/")
Dim s
FileName=GetCurrentHost & strUPLOADDIR &"\" & upload.form("edtFileLoad_Name")
If upload.Save("edtFileLoad",0)=True Then
	Dim uf
	Set uf=New TUpLoadFile
	uf.AuthorID=BlogUser.ID
	uf.AutoName=False
	uf.IsManual=True
	uf.FileSize=upload.form("edtFileLoad_Size")
	uf.FileName=upload.form("edtFileLoad_Name")
	uf.UpLoad
End If

Dim strJSON
strJSON="{'state':'"& upload.Error2Info("edtFileLoad") & "','url':'"& upload.form("edtFileLoad") &"','fileType':'"&upload.form("edtFileLoad_Ext")&"','title':'"&TransferHTML(upload.form("pictitle"),"[&][<][>][""][space][enter][nohtml]")&"','original':'"&upload.Form("edtFileLoad_Name")&"'}"

	
For Each sAction_Plugin_uEditor_FileUpload_End in Action_Plugin_uEditor_FileUpload_End
	If Not IsEmpty(sAction_Plugin_uEditor_FileUpload_End) Then Call Execute(sAction_Plugin_uEditor_FileUpload_End)
Next
response.AddHeader "json",strjson
response.write strJSON

set upload=nothing
Call System_Terminate()

%>