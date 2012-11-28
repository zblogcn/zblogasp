<%
''*****************************************************
'   ZSXSOFT 数据库操作类
''*****************************************************
Class QQConnect_DB
	Dim objRS
	Public ID
	Public objUser
	Public OpenID
	Public AccessToken
	
	Sub Class_Initialize()  '初始化类并创建数据库
		Set objUser=New TUser
		'On Error Resume Next
		'objConn.Execute "SELECT TOP 1 [QQ_ID] FROM [blog_Plugin_QQConnect] "
		'If Err.Number<>0 Then
		'	Call CreateDB
		'	Err.Clear
		'End If
	End Sub
	
	Public Sub CreateDB() '创建数据库
		IF ZC_MSSQL_ENABLE=True Then
			objConn.execute("CREATE TABLE [blog_Plugin_QQConnect] (QQ_ID int identity(1,1) not null primary key,QQ_UserID int default 0,QQ_OpenID nvarchar(32) default '',QQ_AToken nvarchar(32) default '')")
		Else
			objConn.execute("CREATE TABLE [blog_Plugin_QQConnect] (QQ_ID AutoIncrement primary key,QQ_UserID int default 0,QQ_OpenID VARCHAR(32) default """",QQ_AToken VARCHAR(32) default """")")
		End If
	End Sub
	
	Function LoadInfo(Typ) '读取用户信息，使用ID、OPENID、EMAIL、用户ID来读取，同时兼备判断是否存在功能
		LoadInfo=False
		Dim strSQL
		strSQL="SELECT [QQ_ID],[QQ_UserID],[QQ_OpenID],[QQ_AToken] FROM [blog_Plugin_QQConnect] WHERE "
		Select Case Typ
			Case 1,1000
				Call CheckParameter(ID,"int",0)
				strSQL=strSQL & "QQ_ID="&ID
			Case 2,2000
				Call CheckParameter(objUser.ID,"int",0)
				strSQL=strSQL & "QQ_USERID="&objUser.ID
			Case 4,5,4000
				If CheckRegExp(OpenID,"^[0-9A-Z]{32}$") Then
					OpenID=FilterSQL(OpenID)
					strSQL=strSQL & "QQ_OpenID='"&OpenID&"'"
				Else
					LoadInfo=False
					Exit Function
				End If
		End Select
		Set objRS=objConn.Execute(strSQL)
		If (Not objRS.bof) And (Not objRS.eof) Then
			If Typ<1000 Then
				ID=objRS("QQ_ID")
				If Typ<>5 Then objUser.LoadInfoById CInt(objRS("QQ_UserID"))
				OpenID=objRS("QQ_OpenID")
				AccessToken=objRs("QQ_AToken")
			End If
			LoadInfo=True
		End If
		objRS.Close
		Set objRS=Nothing
	End Function

	Function Del()  '删除某个ID的绑定
		Call CheckParameter(ID,"int",0)
		If ID=0 Then Exit Function
		objConn.Execute "DELETE FROM [blog_Plugin_QQConnect] WHERE [QQ_ID]="&ID
	End Function

	Function Bind()   '将数据库里OpenID与现有帐号绑定
		Dim strSQL
		Call CheckParameter(objUser.ID,"int",0)
		OpenID=FilterSQL(OpenID)
		AccessToken=FilterSQL(AccessToken)
		'If objUser.ID=0 Then
		'	Bind=False
	'	'	Exit Function
		'End If
		If Not CheckRegExp(OpenID,"^[0-9A-Z]{32}$") Then
			Call ShowError(3)
		End If
		If OpenID="" Or AccessToken="" Then Bind=False:Exit Function
		If LoadInfo(4000) Then
			strSQL="UPDATE [blog_Plugin_QQConnect] SET [QQ_UserID]="&objUser.ID&",[QQ_OpenID]='"&OpenID&"',[QQ_AToken]='"&AccessToken&"' WHERE [QQ_OpenID]='"&OpenID&"'"
		Else
			strSQL="INSERT INTO [blog_Plugin_QQConnect] ([QQ_UserID],[QQ_OpenID],[QQ_AToken]) VALUES ("&objUser.ID&",'"&OpenID&"','"&AccessToken&"')"
		End If
		'response.write strsql
		objConn.Execute strSQL
		Dim objRS
		Set objRS=objConn.Execute("SELECT MAX([QQ_ID]) FROM [blog_Plugin_QQConnect]")
		If (Not objRS.bof) And (Not objRS.eof) Then
			ID=objRS(0)
		End If
		Set objRS=Nothing
	End Function
	
	Function Login() '用QQ登录
		LoadInfo 4
		BlogUser.LoginType="Self"
		BlogUser.Name=objUser.name
		BlogUser.PassWord=objUser.Password
		If BlogUser.Verify=True Then
			Response.Cookies("password")=BlogUser.PassWord
			If Request.Form("savedate")<>0 Then
				Response.Cookies("password").Expires = DateAdd("d", 1, now)
			End If
			Response.Cookies("password").Path = CookiesPath()
			Login=True
		End If
		Response.Cookies("username")=escape(BlogUser.name)
		If Request.Form("savedate")<>0 Then
			Response.Cookies("username").Expires = DateAdd("d", 1, now)
		End If
		Response.Cookies("username").Path = CookiesPath()
	End Function
End Class
%>