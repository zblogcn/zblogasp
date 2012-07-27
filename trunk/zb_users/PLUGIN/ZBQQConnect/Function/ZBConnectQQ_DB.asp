<%
Class ZBConnectQQ_DB
	Dim objRS
	Public ID
	Public objUser
	Public OpenID
	Public AccessToken
	Public Head
			
	Private EmailMD5
	
	Public Property Let EMail(str)
		EMailMD5=str
	End Property
	Public Property Get EMail
		EMail=EMailMD5
	End Property
	
	
	Sub Class_Initialize()
		Set objUser=New TUser
	End Sub
	
	Sub CreateDB()
		IF ZC_MSSQL=True Then
			objConn.execute("CREATE TABLE [blog_Plugin_ZBQQConnect] (QQ_ID int identity(1,1) not null primary key,QQ_UserID int default 0,QQ_EmlMD5 nvarchar(32) default '',QQ_OpenID nvarchar(32) default '',QQ_AToken nvarchar(32) default '',QQ_Head nvarchar(32) default '')")
		Else
			objConn.execute("CREATE TABLE [blog_Plugin_ZBQQConnect] (QQ_ID AutoIncrement primary key,QQ_UserID int default 0,QQ_EmlMD5 VARCHAR(32) default """",QQ_OpenID VARCHAR(32) default """",QQ_AToken VARCHAR(32) default """",QQ_Head VARCHAR(32) default """"")
		End If
	End Sub
	
	Function LoadInfo(Typ)
		Dim strSQL
		strSQL="SELECT [QQ_ID],[QQ_UserID],[QQ_EmlMD5],[QQ_OpenID],[QQ_AToken],[QQ_Head] FROM [blog_Plugin_ZBQQConnect] WHERE "
		Select Case Typ
			Case 1
				Call CheckParameter(ID,"int",0)
				strSQL=strSQL & "QQ_ID="&ID
			Case 2
				Call CheckParameter(objUser.ID,"int",0)
				strSQL=strSQL & "QQ_USERID="&objUser.ID
			Case 3
				strSQL=strSQL & "QQ_EmlMD5='"&EMailMD5&"'"
			Case 4
				strSQL=strSQL & "QQ_OpenID='"&OpenID&"'"
		End Select
		Set objRS=objConn.Execute(strSQL)
		If (Not objRS.bof) And (Not objRS.eof) Then
			ID=objRS("QQ_ID")
			objUser.LoadInfoById CInt(objRS("QQ_UserID"))
			EmailMD5=objRs("QQ_EmlMD5")
			OpenID=objRS("QQ_OpenID")
			AccessToken=objRs("QQ_AToken")
			Head=objRs("QQ_Head")
			LoadInfo=True
		End If
		objRS.Close
		Set objRS=Nothing
	End Function


	Function Bind()
		Dim strSQL
		'Call CheckParameter(ID,"int",0)
		Call CheckParameter(objUser.ID,"int",0)
		OpenID=FilterSQL(OpenID)
		AccessToken=FilterSQL(AccessToken)
		Head=FilterSQL(Head)
		EmailMD5=LCase(FilterSQL(EmailMD5))
		If objUser.ID=0 And Len(EmailMD5)<>32 Then
			Bind=False
			Exit Function
		ElseIf objUser.ID>0 And Len(EmailMD5)<>32 Then
			objUser.LoadInfoById objUser.ID
			EmailMD5=MD5(objUser.EMail)
		End If
		If OpenID="" Or AccessToken="" Then Bind=False:Exit Function
		If LoadInfo(3) Then
			strSQL="UPDATE [blog_Plugin_ZBQQConnect] SET [QQ_UserID]="&objUser.ID&",[QQ_OpenID]='"&OpenID&"',[QQ_AToken]='"&AccessToken&"',[QQ_Head]='"& Head&"' WHERE [QQ_EmlMD5]='"&EmailMD5&"'"
		Else
			strSQL="INSERT INTO [blog_Plugin_ZBQQConnect] ([QQ_UserID],[QQ_OpenID],[QQ_AToken],[QQ_Head],[QQ_EmlMD5]) VALUES ("&objUser.ID&",'"&OpenID&"','"&AccessToken&"','"& Head&"','"&EmailMD5&"')"
		End If
		objConn.Execute strSQL
		Dim objRS
		Set objRS=objConn.Execute("SELECT MAX([QQ_ID]) FROM [blog_Plugin_ZBQQConnect]")
		If (Not objRS.bof) And (Not objRS.eof) Then
			ID=objRS(0)
		End If
		Set objRS=Nothing
	End Function
	
	
	Function SetHead()
		objConn.Execute "UPDATE [blog_Plugin_ZBQQConnect] SET [QQ_Head]='"&FilterSQL(Head)&"'"
	End Function
	
	Function Del()
		objConn.Execute "DELETE FROM [blog_Plugin_ZBQQConnect] WHERE [QQ_ID]="&ID
	End Function
	
	Function Login()
		LoadInfo 4
		BlogUser.LoginType="Self"
		BlogUser.Name=objUser.name
		BlogUser.PassWord=objUser.Password

		If BlogUser.Verify=True Then

			Response.Cookies("password")=BlogUser.PassWord
			If Request.Form("savedate")<>0 Then
				Response.Cookies("password").Expires = DateAdd("d", 1, now)
			End If
			Response.Cookies("password").Path = "/"
			Login=True

		End If

		Response.Cookies("username")=escape(BlogUser.name)
		If Request.Form("savedate")<>0 Then
			Response.Cookies("username").Expires = DateAdd("d", 1, now)
		End If
		Response.Cookies("username").Path = "/"

	End Function
End Class
%>