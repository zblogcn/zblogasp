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
Dim objUploadFile
Set objUpLoadFile=New TUpLoadFile
objUpLoadFile.AuthorID=BlogUser.ID
Dim state
state="SUCCESS" 
Dim strFileType
Dim strFileName
Dim strUPLOADDIR
Dim strUPLOADDIR2,bolIsRen
If objUpLoadFile.UpLoad(True) Then
	If ZC_UPLOAD_DIRBYMONTH Then
			CreatDirectoryByCustomDirectory(ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now())))
			strUPLOADDIR = ZC_UPLOAD_DIRECTORY&"/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now())) & "/"
			strUPLOADDIR2 = "zb_users/upload/"&Year(GetTime(Now()))&"/"&Month(GetTime(Now())) & "/"
	Else
			strUPLOADDIR = ZC_UPLOAD_DIRECTORY & "/"
			strUPLOADDIR2 ="zb_users/upload/"
	End If
	strFileType=LCase(objUpLoadFile.FileName)
	strFileType=Split(strFileType,".")(Ubound(Split(strFileType,".")))
	If err.number<>0 Then state=err.description
	response.Write "{'state':'"& state & "','url':'"& objUpLoadFile.FileName &"','fileType':'"&strFileType&"'}"
End If	

	
For Each sAction_Plugin_uEditor_FileUpload_End in Action_Plugin_uEditor_FileUpload_End
	If Not IsEmpty(sAction_Plugin_uEditor_FileUpload_End) Then Call Execute(sAction_Plugin_uEditor_FileUpload_End)
Next

set upload=nothing
Call System_Terminate()
    function htmlspecialchars(someString)
        htmlspecialchars = replace(replace(replace(replace(someString, "&", "&amp;"), ">", "&gt;"), "<", "&lt;"), """", "&quot;")
    end function
%>