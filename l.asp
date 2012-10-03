<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="zb_users/c_option.asp" -->
<!-- #include file="zb_system/function/c_function.asp" -->
<!-- #include file="zb_system/function/c_system_lib.asp" -->
<!-- #include file="zb_system/function/c_system_base.asp" -->
<!-- #include file="zb_system/function/c_system_event.asp" -->
<!-- #include file="zb_system/function/c_system_plugin.asp" -->
<!-- #include file="zb_users/plugin/p_config.asp" -->
<%
OpenConnect()

Function RevToComment()

	ShowError_Custom="StarTime = Timer()"

	Call GetUser()

	Dim objRS2,objComment,s,t,Match,Matches,t1,t2,t3,t4,c,u

	Dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase =True
	objRegExp.Global=True


	Set objRS2=objConn.Execute("SELECT * FROM [blog_Comment] WHERE ([comm_isCheck]=0) AND (InStr(1,LCase([comm_Content]),LCase('[/REVERT]'),0)<>0)")
	If (Not objRS2.bof) And (Not objRS2.eof) Then
		Do While Not objRS2.eof

			Set objComment=New TComment
			objComment.LoadInfoByArray(Array(objRS2("comm_ID"),objRS2("log_ID"),objRS2("comm_AuthorID"),objRS2("comm_Author"),objRS2("comm_Content"),objRS2("comm_Email"),objRS2("comm_HomePage"),objRS2("comm_PostTime"),objRS2("comm_IP"),objRS2("comm_Agent"),objRS2("comm_Reply"),objRS2("comm_LastReplyIP"),objRS2("comm_LastReplyTime"),objRS2("comm_ParentID"),objRS2("comm_IsCheck"),objRS2("comm_Meta")))



			s=objComment.Content
			objRegExp.Pattern="(\[REVERT=)(.+?)(\])([\u0000-\uffff]+?)(\[\/REVERT\])"
			Set Matches = objRegExp.Execute(s)

			For Each Match in Matches

				t=Match
				s=Replace(s,t,"")


				t1=Match.SubMatches(1)
				t2=Match.SubMatches(3)

				t4= Left(t1,InStr(1,t1," ",0)-1)
				t3=Mid(t1,InStr(InStr(1,t1," ",0)+1,t1," ",0),InstrRev(t1," ",Len(t1),0)-InStr(InStr(1,t1," ",0)+1,t1," ",0) )

				t3=CDate(t3)

				u=0
				Dim User
				For Each User in Users
					If IsObject(User) Then
						If User.Name=t4 Then
							u=User.ID
						End If
					End If
				Next

				Set c=new TComment

				c.log_ID=objComment.log_ID
				c.AuthorID=u
				c.Author=t4
				c.Content=t2
				c.Email=""
				c.HomePage=""
				c.ParentID=objComment.ID
				c.PostTime=t3

				c.post
				Set c=Nothing


			Next
			Set Matches = Nothing



			objComment.Content=s

			objConn.Execute("UPDATE [blog_Comment] SET [comm_Content]='"&FilterSQL(s)&"' WHERE [comm_ID] =" & objComment.ID)

			Set objComment=Nothing
			objRS2.MoveNext
		Loop
	End If
	objRS2.Close
	Set objRS2=Nothing

End Function




Call  RevToComment()

response.write "ok"
%>