<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件应用:    1.9 其它版本的Z-blog未知
'// 插件制作:    ZSXSOFT(http://www.zsxsoft.com/)
'// 备    注:    Gravatar - 挂口函数页
'///////////////////////////////////////////////////////////////////////////////

'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************
Dim Gravatar_EmailMD5
Dim Gravatar_Enable
Dim Gravatar_Refresh

'注册插件
Call RegisterPlugin("Gravatar","ActivePlugin_Gravatar")
'挂口部分
Function ActivePlugin_Gravatar()	

	Gravatar_Refresh=False

	Dim c
	Set c=New TConfig
	c.Load "Gravatar"
	Gravatar_EmailMD5=c.Read("c")
	Gravatar_Enable=CBool(c.Read("e"))

	If Gravatar_Enable=True Then
		Call Add_Action_Plugin("Action_Plugin_TComment_Avatar","If FAvatar="""" Then FAvatar=Gravatar_Add(AuthorID,EmailMD5)")
	End If

	Call Add_Response_Plugin("Response_Plugin_EditUser_Form","<label><input name='gravatar_cacheimage' type='checkbox' value='gravatar'>&nbsp;&nbsp;刷新用户Gravatar头像并缓存在Blog里.</label>")

	Call Add_Filter_Plugin("Filter_Plugin_EditUser_Succeed","Gravatar_Filter_Plugin_EditUser_Succeed")

End Function



Function Gravatar_Add(AuthorID,EmailMD5)

	If AuthorID>0 Then
	  Dim fso
	  Set fso = CreateObject("Scripting.FileSystemObject")
	  If (fso.FileExists(BlogPath & "zb_users/avatar/"&AuthorID&".png")) Then
		Gravatar_Add=GetCurrentHost() & "zb_users/avatar/"&AuthorID&".png"
	  Else
		Gravatar_Add=Replace(Replace(Gravatar_EmailMD5,"{%emailmd5%}",EmailMD5),"{%source%}",Server.URLEncode(ZC_BLOG_HOST&"zb_users/avatar/0.png"))
	  End If
	Else
		If EmailMD5<>"" Then
			Gravatar_Add=Replace(Replace(Gravatar_EmailMD5,"{%emailmd5%}",EmailMD5),"{%source%}",Server.URLEncode(ZC_BLOG_HOST&"zb_users/avatar/0.png"))
		Else
			Gravatar_Add=GetCurrentHost() & "zb_users/avatar/0.png"
		End If
	End If

End Function

Sub Gravatar_GetImage(ID)

	On Error Resume Next

	Dim k

	k=Replace(Replace(Gravatar_EmailMD5,"{%emailmd5%}",MD5(objConn.Execute("SELECT [mem_Email] FROM [blog_Member] WHERE [mem_ID]="&ID)(0))),"{%source%}",Server.URLEncode(ZC_BLOG_HOST&"zb_users/avatar/0.png"))

	dim u,v,w
	set u=server.createobject("msxml2.serverxmlhttp")
	u.open "GET",k
	u.send
	If u.Readystate =4 Then
		If u.status=200 Then 
			v=u.ResponseBody 
			set w=server.createObject("Adodb.Stream") 
			w.Type = 1 
			w.Open 
			w.Write v 
			w.SaveToFile BlogPath & "zb_users/avatar/"&ID&".png",2 
			w.Close() 
		End If
	End If
	Set w=nothing 
	Set u=nothing

	Err.Clear

End Sub



Function InstallPlugin_Gravatar
	Dim c
	Set c=New TConfig
	c.Load "Gravatar"
	c.Write "v","1.0"
	c.Write "e","True"
	c.Write "c","http://cn.gravatar.com/avatar/{%emailmd5%}?s=40&d={%source%}"
	c.Save
End Function


Function Gravatar_Filter_Plugin_EditUser_Succeed(objUser)

If Request.Form("gravatar_cacheimage")<>"" Then

Call AddBatch("缓存用户"& objUser.FirstName&"的Gravatar头像","Gravatar_GetImage "& objUser.ID)

End IF

End Function

%>