<%@ CODEPAGE=65001 %>
<!--#include file="ASPIncludeFile.asp"-->
<%
uEditor_i

For Each sAction_Plugin_uEditor_imageManager_Begin in Action_Plugin_uEditor_imageManager_Begin
	If Not IsEmpty(sAction_Plugin_uEditor_imageManager_Begin) Then Call Execute(sAction_Plugin_uEditor_imageManager_Begin)
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
			If CheckRegExp(objRS("ul_FileName"),"\.(jpe?g|gif|bmp|png)$")=True Then
				If IsNull(objRS("ul_DirByTime"))=False And objRS("ul_DirByTime")<>"" Then
					If CBool(objRS("ul_DirByTime"))=True Then
						Response.Write ZC_UPLOAD_DIRECTORY &"/"&Year(objRS("ul_PostTime")) & "/" & Month(objRS("ul_PostTime")) & "/"&objRS("ul_FileName")&"ue_separate_ue"
					Else
						Response.Write ZC_UPLOAD_DIRECTORY &"/"&objRS("ul_FileName")&uEditor_Split
					End If
				Else
					Response.Write ZC_UPLOAD_DIRECTORY &"/"&objRS("ul_FileName")&uEditor_Split
				End If
			End If
			objRS.MoveNext
		Next
	End If
For Each sAction_Plugin_uEditor_imageManager_End in Action_Plugin_uEditor_imageManager_End
	If Not IsEmpty(sAction_Plugin_uEditor_imageManager_End) Then Call Execute(sAction_Plugin_uEditor_imageManager_End)
Next
	Response.Write strResponse
Call System_Terminate()
%>