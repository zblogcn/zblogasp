<!--#include file="up_inc.asp"-->
<!-- #include file="../../../../zb_users\c_option.asp" -->
<!-- #include file="../../../function\c_function.asp" -->
<!-- #include file="../../../function\c_function_md5.asp" -->
<!-- #include file="../../../function\c_system_lib.asp" -->
<!-- #include file="../../../function\c_system_base.asp" -->
<!-- #include file="../../../function\c_system_event.asp" -->
<!-- #include file="../../../function\c_system_plugin.asp" -->
<!-- #include file="../../../function\rss_lib.asp" -->
<!-- #include file="../../../function\atom_lib.asp" -->
<!-- #include file="../../../../zb_users\plugin\p_config.asp" -->
<%
On Error Resume Next
Call System_Initialize()
Call CheckReference("")
If Not CheckRights("ArticleEdt") Then Call ShowError(6)

For Each sAction_Plugin_imageManager_Begin in Action_Plugin_imageManager_Begin
	If Not IsEmpty(sAction_Plugin_imageManager_Begin) Then Call Execute(sAction_Plugin_imageManager_Begin)
Next
	Dim strResponse,objUpload,objRS,intPageAll
	If CheckRights("Root")=False Then strSQL="WHERE ([ul_AuthorID] = " & BlogUser.ID & ")"
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	objRS.Source=""
	objRS.Open("SELECT * FROM [blog_UpLoad] " & strSQL & " ORDER BY [ul_PostTime] DESC")
	objRS.PageSize=ZC_MANAGE_COUNT
	Call CheckParameter(intPage,"int",1)
	If objRS.PageCount>0 Then objRS.AbsolutePage = intPage
	intPageAll=objRS.PageCount

	If (Not objRS.bof) And (Not objRS.eof) Then

		For i=1 to objRS.PageSize
			If CheckRegExp(objRS("ul_FileName"),".+?\.jp[e]?g|.+?\.gif|.+?\.png|.+?\.bmp|.+?\.tif[f]?")=True Then
				If IsNull(objRS("ul_DirByTime"))=False And objRS("ul_DirByTime")<>"" Then
					If CBool(objRS("ul_DirByTime"))=True Then
						Response.Write "../../../../zb_users/"& ZC_UPLOAD_DIRECTORY &"/"&Year(objRS("ul_PostTime")) & "/" & Month(objRS("ul_PostTime")) & "/"&objRS("ul_FileName")&"ue_separate_ue"
					Else
						Response.Write "../../../../zb_users/"& ZC_UPLOAD_DIRECTORY &"/"&objRS("ul_FileName")&"ue_separate_ue"
					End If
				Else
					Response.Write "../../../../zb_users/"& ZC_UPLOAD_DIRECTORY &"/"&objRS("ul_FileName")&"'ue_separate_ue"
				End If
			End If
			objRS.MoveNext
		Next
	End If
For Each sAction_Plugin_imageManager_End in Action_Plugin_imageManager_End
	If Not IsEmpty(sAction_Plugin_imageManager_End) Then Call Execute(sAction_Plugin_imageManager_End)
Next
	Response.Write strResponse
Call System_Terminate()
%>