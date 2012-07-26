<%
Class ZBConnectQQ_DB
	Dim objRS
	Public ID
	Public objUser
	Public OpenID
	Public AccessToken
	Public Head
	
	Sub Class_Initialize()
		Set objUser=New TUser
	End Sub
	
	Sub CreateDB()
		IF ZC_MSSQL=True Then
			objConn.execute("CREATE TABLE [blog_Plugin_ZBQQConnect] (QQ_ID int identity(1,1) not null primary key,QQ_UserID int default 0,QQ_OpenID nvarchar(32) default '',QQ_AToken nvarchar(32) default '',QQ_Head nvarchar(32) default '')")
		Else
			objConn.execute("CREATE TABLE [blog_Plugin_ZBQQConnect] (QQ_ID AutoIncrement primary key,QQ_UserID int default 0,QQ_OpenID VARCHAR(32) default """",QQ_AToken VARCHAR(32) default """",QQ_Head VARCHAR(32) default """"")
		End If
	End Sub
	
	Function LoadInfoByUserID(ID)
		Call CheckParameter(ID,"int",0)
		Set objRS=objConn.Execute("SELECT [QQ_ID],[QQ_UserID],[QQ_OpenID],[QQ_AToken],[QQ_Head] FROM [blog_Plugin_ZBQQConnect] WHERE QQ_USERID="&ID)
		If (Not objRS.bof) And (Not objRS.eof) Then
			ID=objRS("QQ_ID")
			objUser.LoadInfoById(objRS("QQ_UserID"))
			OpenID=objRS("QQ_OpenID")
			AccessToken=objRs("QQ_AToken")
			Head=objRs("QQ_Head")
			LoadInfoByID=True
		End If
		objRS.Close
		Set objRS=Nothing
	End Function
	Function LoadInfoByID(ID)
		Call CheckParameter(ID,"int",0)
		Set objRS=objConn.Execute("SELECT [QQ_ID],[QQ_UserID],[QQ_OpenID],[QQ_AToken],[QQ_Head] FROM [blog_Plugin_ZBQQConnect] WHERE QQ_ID="&ID)
		If (Not objRS.bof) And (Not objRS.eof) Then
			ID=objRS("QQ_ID")
			objUser.LoadInfoById(objRS("QQ_UserID"))
			OpenID=objRS("QQ_OpenID")
			AccessToken=objRs("QQ_AToken")
			Head=objRs("QQ_Head")
			LoadInfoByID=True
		End If
		objRS.Close
		Set objRS=Nothing
	End Function
End Class
%>